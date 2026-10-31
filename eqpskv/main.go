// Command eqpskv builds an SKV 2198 table ("Genomsnittligt omkostnadsbelopp") and a K4
// sales summary for Saab B shares acquired through EquatePlus (Saab Share Matching Plan)
// and sold through Nordnet.
//
// Only the Go standard library is used.
//
// Input files (looked up in the given directory):
//
//	PortfolioDetails*.xlsx              all allocations: purchases, vested matching shares,
//	                                    free shares (primary source of acquisitions)
//	CompletedTransactions*.xlsx         deliveries (Leverans), transfers out (Överför) and
//	                                    dividends (used as holding checkpoints)
//	Sammanfattning av inköpscykel*.csv  monthly purchases (detects allocations that
//	                                    EquatePlus restated after the 2024 split)
//	Nordnet*.csv                        Nordnet "transaktioner och notor" exports (one or
//	                                    more, UTF-16, tab separated); SAAB B, old and new ISIN
//
// Outputs:
//
//	skv2198_saab.csv  SKV 2198 columns 1-8 (+ comment), chronological
//	k4_saab.csv       one row per sale and per-year totals for K4 section A
//	k4_innehav_saab.csv  holdings log: one row per event (purchases, deliveries, matching
//	                  grants and vestings, transfers, sales, split) with K4 values on sale
//	                  rows and holdings after each event in EquatePlus (free for trading /
//	                  with restrictions / unvested company match), under transfer, in
//	                  Nordnet and in total (units valid on the date; unvested matching is
//	                  not part of the total)
//
// Tax treatment used:
//   - Purchased Shares (EquatePlus) and KÖPT (Nordnet) -> "Köp", cost incl. fees.
//   - Matching / Free shares -> "Förm" (förmånsbeskattat belopp): market value at delivery.
//   - Saab 4:1 split -> "S": quantity x4, total cost unchanged.
//   - SÅLT (Nordnet) -> "Förs": utnyttjat omkostnadsbelopp = antal x genomsnitt.
//   - OMVANDLING INLÄGG VP (Nordnet) -> "Övr" with Nordnet's inköpsvärde.
//   - Transfers EquatePlus -> Nordnet are not disposals and do not appear in the table.
//   - Unvested matching shares (begränsade aktieenheter) are excluded.
//   - If the Nordnet export does not reach back far enough to explain where all shares
//     transferred out of EquatePlus went, an adjustment row ("Övr") removes the
//     unexplained shares at the average cost at the start of the export, with a warning.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/csv"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"
)

const (
	splitRatio         = 4
	securityPrefix     = "SAAB B"
	tol                = 0.001
	transferWindowDays = 60
	holdingYears       = 3  // Share Matching Plan holding period for purchased shares
	deliveryWindowDays = 45 // max distance between portfolio vesting date and delivery
)

// Saab AB 4:1 share split (spring 2024). Taken from the Nordnet "SPLIT INLÄGG VP" row
// when present; otherwise this fallback (override with -split-date).
var fallbackSplitDate = mkDate(2024, 4, 19)

var svMonths = map[string]time.Month{
	"jan": 1, "feb": 2, "mars": 3, "mar": 3, "apr": 4, "maj": 5, "juni": 6, "jun": 6,
	"juli": 7, "jul": 7, "aug": 8, "sep": 9, "sept": 9, "okt": 10, "nov": 11, "dec": 12,
}

var warnings []string

func warn(format string, a ...any) { warnings = append(warnings, fmt.Sprintf(format, a...)) }

func fatal(format string, a ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", a...)
	os.Exit(1)
}

// ---------------------------------------------------------------------------- dates

// Date is a calendar date (UTC midnight).
type Date struct{ time.Time }

func mkDate(y int, m time.Month, d int) Date { return Date{time.Date(y, m, d, 0, 0, 0, 0, time.UTC)} }

func (d Date) String() string       { return d.Format("2006-01-02") }
func (d Date) IsZero() bool         { return d.Time.IsZero() }
func (d Date) Before(o Date) bool   { return d.Time.Before(o.Time) }
func (d Date) After(o Date) bool    { return d.Time.After(o.Time) }
func (d Date) AddDays(n int) Date   { return Date{d.Time.AddDate(0, 0, n)} }
func (d Date) DaysUntil(o Date) int { return int(o.Time.Sub(d.Time).Hours() / 24) }
func (d Date) Year4() string        { return d.Format("2006") }
func (d Date) Cmp(o Date) int       { return d.Time.Compare(o.Time) }

func parseISO(s string) (Date, error) {
	t, err := time.Parse("2006-01-02", strings.TrimSpace(s))
	return Date{t}, err
}

// addYears adds whole years; 29 Feb becomes 28 Feb in non-leap years (as in the Python version).
func addYears(d Date, years int) Date {
	y := d.Year() + years
	day := d.Day()
	if d.Month() == time.February && day == 29 && !(y%4 == 0 && (y%100 != 0 || y%400 == 0)) {
		day = 28
	}
	return mkDate(y, d.Month(), day)
}

func excelDate(v string) Date {
	f, err := strconv.ParseFloat(strings.TrimSpace(v), 64)
	if err != nil {
		fatal("ogiltigt Excel-datum %q", v)
	}
	return mkDate(1899, 12, 30).AddDays(int(f))
}

// svDate parses "16 nov. 2023" / "16 maj 2024".
func svDate(s string) Date {
	parts := strings.Fields(strings.ReplaceAll(s, ".", ""))
	if len(parts) != 3 {
		fatal("ogiltigt datum %q", s)
	}
	d, err1 := strconv.Atoi(parts[0])
	y, err2 := strconv.Atoi(parts[2])
	m, ok := svMonths[strings.ToLower(parts[1])]
	if err1 != nil || err2 != nil || !ok {
		fatal("ogiltigt datum %q", s)
	}
	return mkDate(y, m, d)
}

// ---------------------------------------------------------------------------- numbers

// num parses "1 234,56", "1234.56", "-" or "" (the latter two give ok=false).
func num(s string) (float64, bool) {
	s = strings.TrimSpace(s)
	s = strings.ReplaceAll(s, " ", "")
	s = strings.ReplaceAll(s, " ", "")
	if s == "" || s == "-" {
		return 0, false
	}
	f, err := strconv.ParseFloat(strings.ReplaceAll(s, ",", "."), 64)
	if err != nil {
		return 0, false
	}
	return f, true
}

func numOr0(s string) float64 { f, _ := num(s); return f }

func mustFloat(s string) float64 {
	f, ok := num(s)
	if !ok {
		fatal("ogiltigt tal %q", s)
	}
	return f
}

// g formats like Python's "{:g}" for typical values.
func g(f float64) string { return strconv.FormatFloat(f, 'g', -1, 64) }

// fmtSv formats with thousands separator ' ' and decimal comma.
func fmtSv(v float64, decimals int) string {
	s := strconv.FormatFloat(v, 'f', decimals, 64)
	neg := strings.HasPrefix(s, "-")
	s = strings.TrimPrefix(s, "-")
	intPart, frac := s, ""
	if i := strings.IndexByte(s, '.'); i >= 0 {
		intPart, frac = s[:i], s[i+1:]
	}
	var b strings.Builder
	for i, c := range intPart {
		if i > 0 && (len(intPart)-i)%3 == 0 {
			b.WriteByte(' ')
		}
		b.WriteRune(c)
	}
	out := b.String()
	if frac != "" {
		out += "," + frac
	}
	if neg && strings.Trim(out, "0, ") != "" {
		out = "-" + out
	}
	return out
}

// ---------------------------------------------------------------------------- xlsx

type xSST struct {
	SI []struct {
		T string `xml:"t"`
		R []struct {
			T string `xml:"t"`
		} `xml:"r"`
	} `xml:"si"`
}

type xSheet struct {
	Rows []struct {
		C []struct {
			R  string `xml:"r,attr"`
			T  string `xml:"t,attr"`
			V  string `xml:"v"`
			IS struct {
				T string `xml:"t"`
				R []struct {
					T string `xml:"t"`
				} `xml:"r"`
			} `xml:"is"`
		} `xml:"c"`
	} `xml:"sheetData>row"`
}

func zipRead(z *zip.ReadCloser, name string) ([]byte, bool) {
	for _, f := range z.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				fatal("%s: %v", name, err)
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			if err != nil {
				fatal("%s: %v", name, err)
			}
			return b, true
		}
	}
	return nil, false
}

// readXLSX returns the rows of the first sheet as {column letter: value} maps.
func readXLSX(path string) []map[string]string {
	z, err := zip.OpenReader(path)
	if err != nil {
		fatal("%s: %v", path, err)
	}
	defer z.Close()
	var shared []string
	if b, ok := zipRead(z, "xl/sharedStrings.xml"); ok {
		var sst xSST
		if err := xml.Unmarshal(b, &sst); err != nil {
			fatal("%s: %v", path, err)
		}
		for _, si := range sst.SI {
			s := si.T
			for _, r := range si.R {
				s += r.T
			}
			shared = append(shared, s)
		}
	}
	b, ok := zipRead(z, "xl/worksheets/sheet1.xml")
	if !ok {
		fatal("%s: sheet1 saknas", path)
	}
	var sh xSheet
	if err := xml.Unmarshal(b, &sh); err != nil {
		fatal("%s: %v", path, err)
	}
	var rows []map[string]string
	for _, r := range sh.Rows {
		row := map[string]string{}
		for _, c := range r.C {
			col := strings.TrimRightFunc(c.R, unicode.IsDigit)
			var val string
			switch c.T {
			case "s":
				if i, err := strconv.Atoi(strings.TrimSpace(c.V)); err == nil && i < len(shared) {
					val = shared[i]
				}
			case "inlineStr":
				val = c.IS.T
				for _, rr := range c.IS.R {
					val += rr.T
				}
			default:
				val = c.V
			}
			row[col] = val
		}
		rows = append(rows, row)
	}
	return rows
}

// table finds the header row whose column A == headerKey and returns the following
// non-empty rows keyed by header name.
func table(rows []map[string]string, headerKey string) []map[string]string {
	for i, r := range rows {
		if strings.TrimSpace(r["A"]) != headerKey {
			continue
		}
		var out []map[string]string
		for _, rr := range rows[i+1:] {
			empty := true
			for _, v := range rr {
				if strings.TrimSpace(v) != "" {
					empty = false
					break
				}
			}
			if empty {
				continue
			}
			m := map[string]string{}
			for col, name := range r {
				m[name] = rr[col]
			}
			out = append(out, m)
		}
		return out
	}
	fatal("rubrik '%s' hittades inte", headerKey)
	return nil
}

// ---------------------------------------------------------------------------- CSV

func readSvCSV(path string) []map[string]string {
	b, err := os.ReadFile(path)
	if err != nil {
		fatal("%s: %v", path, err)
	}
	b = bytes.TrimPrefix(b, []byte("\xef\xbb\xbf"))
	r := csv.NewReader(bytes.NewReader(b))
	r.Comma = ';'
	r.LazyQuotes = true
	r.FieldsPerRecord = -1
	recs, err := r.ReadAll()
	if err != nil {
		fatal("%s: %v", path, err)
	}
	return recordsToMaps(recs)
}

func recordsToMaps(recs [][]string) []map[string]string {
	if len(recs) == 0 {
		return nil
	}
	header := recs[0]
	var out []map[string]string
	for _, rec := range recs[1:] {
		empty := true
		for _, x := range rec {
			if strings.TrimSpace(x) != "" {
				empty = false
				break
			}
		}
		if empty {
			continue
		}
		m := map[string]string{}
		for i, name := range header {
			name = strings.TrimSpace(name)
			if _, dup := m[name]; dup || i >= len(rec) {
				continue // duplicate "Valuta" columns: keep the first
			}
			m[name] = strings.TrimSpace(rec[i])
		}
		out = append(out, m)
	}
	return out
}

func decodeText(b []byte) string {
	switch {
	case bytes.HasPrefix(b, []byte{0xff, 0xfe}), bytes.HasPrefix(b, []byte{0xfe, 0xff}):
		le := b[0] == 0xff
		b = b[2:]
		u := make([]uint16, len(b)/2)
		for i := range u {
			if le {
				u[i] = uint16(b[2*i]) | uint16(b[2*i+1])<<8
			} else {
				u[i] = uint16(b[2*i])<<8 | uint16(b[2*i+1])
			}
		}
		return string(utf16.Decode(u))
	case bytes.HasPrefix(b, []byte("\xef\xbb\xbf")):
		return string(b[3:])
	case utf8.Valid(b):
		return string(b)
	default: // Latin-1 / cp1252 approximation
		r := make([]rune, len(b))
		for i, c := range b {
			r[i] = rune(c)
		}
		return string(r)
	}
}

func readNordnet(path string) []map[string]string {
	b, err := os.ReadFile(path)
	if err != nil {
		fatal("%s: %v", path, err)
	}
	text := strings.ReplaceAll(decodeText(b), "\r\n", "\n")
	first, _, _ := strings.Cut(text, "\n")
	r := csv.NewReader(strings.NewReader(text))
	r.Comma = ';'
	if strings.Contains(first, "\t") {
		r.Comma = '\t'
	}
	r.LazyQuotes = true
	r.FieldsPerRecord = -1
	recs, err := r.ReadAll()
	if err != nil {
		fatal("%s: %v", path, err)
	}
	return recordsToMaps(recs)
}

func find(dir, pattern string) []string {
	hits, _ := filepath.Glob(filepath.Join(dir, pattern))
	sort.Strings(hits)
	return hits
}

// ---------------------------------------------------------------------------- model

// Event kinds: Köp, Förm, S, Övr, Förs, Adj (adjustment, shown as Övr).
var order = map[string]int{"Adj": -1, "S": 0, "Köp": 1, "Förm": 2, "Övr": 2, "Förs": 3}

type Event struct {
	Date     Date
	Kind     string
	Qty      float64 // shares in (+); for Förs/Adj the number of shares out (positive)
	Cost     float64 // added omkostnadsbelopp (Köp/Förm/Övr)
	Proceeds float64 // Förs: net sale amount after fees
	Note     string
	Ref      string
}

func (e Event) Order() int { return order[e.Kind] }

func sortEvents(ev []Event) {
	sort.SliceStable(ev, func(i, j int) bool {
		if c := ev[i].Date.Cmp(ev[j].Date); c != 0 {
			return c < 0
		}
		return ev[i].Order() < ev[j].Order()
	})
}

type Transfer struct {
	Date Date
	Qty  float64
	Ref  string
}

type Holding struct {
	Date Date
	Qty  float64
}

type Pending struct {
	Alloc, Vest Date
	Qty         float64
}

type EquatePlus struct {
	Events           []Event
	Transfers        []Transfer
	DividendHoldings []Holding
	Pending          []Pending
	Current          float64
	RestrictedLots   []Lot   // purchased shares locked until the matching shares vest
	MatchLots        []Lot   // unvested matching units
	AsOf             Date    // "Per den" date of PortfolioDetails
	RestrictedNow    float64 // Tillgänglig med begränsningar per AsOf
	MatchNow         float64 // unvested matching units per AsOf
}

type Nordnet struct {
	Events      []Event
	Eod         map[Date]float64 // holdings at end of each trade date
	MatchedOut  map[string]bool  // EquatePlus transfer refs matched in Nordnet
	MatchedIn   map[int]bool     // indices into TransfersIn matched to EquatePlus
	TransfersIn []Transfer
	SplitDate   Date
	StartDate   Date
	Opening     float64 // holdings just before StartDate (units of that date)
	Closing     float64
	ClosingDate Date
	Files       []string
}

// ---------------------------------------------------------------------------- EquatePlus

func loadEquatePlus(dir string, splitDate Date, excludeSPA bool) *EquatePlus {
	pf := find(dir, "PortfolioDetails*.xlsx")
	if len(pf) == 0 {
		fatal("PortfolioDetails*.xlsx hittades inte i %s", dir)
	}
	tx := find(dir, "CompletedTransactions*.xlsx")
	cyc := find(dir, "Sammanfattning av inköpscykel*.csv")

	rawPF := readXLSX(pf[0])
	portfolio := table(rawPF, "Allokeringsdatum")
	restatedOn := map[Date]bool{}
	var restrictedLots []Lot
	var transactions []map[string]string
	if len(tx) > 0 {
		transactions = table(readXLSX(tx[0]), "Orderreferens")
	} else {
		warn("CompletedTransactions*.xlsx saknas: inga kontroller mot leveranser/överföringar.")
	}

	boughtOn := map[Date]float64{}
	if len(cyc) > 0 {
		for _, r := range readSvCSV(cyc[0]) {
			if q, ok := num(r["Köpta aktier"]); ok && q != 0 && r["Inköpsdatum"] != "" {
				boughtOn[svDate(r["Inköpsdatum"])] += q
			}
		}
	} else {
		warn("Inköpscykel-CSV saknas: omräkning för split bedöms enbart via intjänandedatum.")
	}

	vestOf := map[Date]Date{}
	for _, r := range portfolio {
		if r["Typ av instrument"] == "begränsade aktieenheter" && r["Tillgänglig från"] != "" {
			vestOf[excelDate(r["Allokeringsdatum"])] = excelDate(r["Tillgänglig från"])
		}
	}

	deliveries := map[Date]float64{}
	for _, t := range transactions {
		if t["Ordertyp"] == "Leverans" && t["Status"] == "Utförda" {
			deliveries[excelDate(t["Datum"])] += mustFloat(t["Antal"])
		}
	}

	var events []Event
	rawPurchasePrice := map[Date]float64{}

	// Purchases.
	for _, r := range portfolio {
		if r["Typ av instrument"] != "aktier" || r["Typ av insättning"] != "Köp" {
			continue
		}
		d := excelDate(r["Allokeringsdatum"])
		qty := mustFloat(r["Allokerad kvantitet"])
		price, ok := num(r["Lösenpris/kostnadsbas"])
		if !ok {
			warn("%s: köp utan pris, hoppas över (%s st)", d, g(qty))
			continue
		}
		restated := false
		if d.Before(splitDate) {
			if csvQ, ok := boughtOn[d]; ok && csvQ != 0 {
				if math.Abs(qty-splitRatio*csvQ) < 0.01 {
					restated = true
				} else if math.Abs(qty-csvQ) > 0.01 {
					warn("%s: portföljens antal %s stämmer inte med inköpscykeln %.5f", d, g(qty), csvQ)
				}
			} else {
				v, ok := vestOf[d]
				if !ok {
					v = d
				}
				restated = !v.Before(splitDate)
			}
		}
		if restated {
			qty, price = qty/splitRatio, price*splitRatio
		}
		rawPurchasePrice[d] = price
		restatedOn[d] = restated
		if v, ok := vestOf[d]; ok { // purchased shares are locked until the matching shares vest
			restrictedLots = append(restrictedLots, Lot{d, v, qty})
		}
		note := "EquatePlus " + r["Instrument"]
		if restated {
			note += " (omräknad till före split)"
		}
		events = append(events, Event{Date: d, Kind: "Köp", Qty: qty, Cost: qty * price, Note: note})
	}

	// Benefit-taxed shares (matching shares at delivery, free shares) and plan account.
	for _, r := range portfolio {
		if r["Typ av instrument"] != "aktier" || r["Typ av insättning"] == "Köp" {
			continue
		}
		d := excelDate(r["Allokeringsdatum"])
		qty := mustFloat(r["Allokerad kvantitet"])
		price, hasPrice := num(r["Lösenpris/kostnadsbas"])
		instr := r["Instrument"]
		if r["Plan"] == "Share Plan Account" {
			msg := fmt.Sprintf("%s: 'Share Plan Account' Tilldelning %s st @ %s kr", d, g(qty), g(price))
			if excludeSPA {
				warn("%s EJ medtagen (-exclude-share-plan-account).", msg)
				continue
			}
			warn("%s medtagen som Övr (behövs för att innehavet ska stämma med "+
				"utdelningarna) - kontrollera omkostnadsbeloppet.", msg)
			note := "EquatePlus " + instr
			// The allocation is sometimes dated after the transfer of the same shares out
			// of EquatePlus; the shares must have been there by the transfer date.
			var earliest Date
			for _, t := range transactions {
				if t["Ordertyp"] != "Överför" || t["Status"] != "Utförda" || math.Abs(mustFloat(t["Antal"])-qty) >= tol {
					continue
				}
				td := excelDate(t["Datum"])
				if days := td.DaysUntil(d); days > 0 && days <= 3 && (earliest.IsZero() || td.Before(earliest)) {
					earliest = td
				}
			}
			if !earliest.IsZero() {
				warn("%s: 'Share Plan Account' daterad till %s då samma antal överfördes ut från EquatePlus det datumet.",
					d, earliest)
				note += fmt.Sprintf(" (daterad till överföringen, tilldelning %s)", d)
				d = earliest
			}
			events = append(events, Event{Date: d, Kind: "Övr", Qty: qty, Cost: qty * price, Note: note})
			continue
		}
		if !hasPrice {
			warn("%s: %s utan pris, hoppas över (%s st)", d, instr, g(qty))
			continue
		}
		note := "EquatePlus " + instr
		if raw, ok := rawPurchasePrice[d]; ok && d.Before(splitDate) {
			ratio := raw / price
			if 0.8*splitRatio < ratio && ratio < 1.2*splitRatio {
				warn("%s: %s pris %s är angivet efter split (köp samma dag %.2f); multiplicerat med %d.",
					d, instr, g(price), raw, splitRatio)
				price *= splitRatio
				note += ", pris justerat för split"
			}
		}
		if dq, ok := deliveries[d]; ok && math.Abs(dq-qty) > tol {
			warn("%s: %s %s st i portföljen men Leverans %s st; leveransens antal används.", d, instr, g(qty), g(dq))
			qty = dq
			note += ", antal enligt Leverans"
		}
		events = append(events, Event{Date: d, Kind: "Förm", Qty: qty, Cost: qty * price, Note: note})
	}

	// Actual delivery date of each matching grant: the Förm event with the same quantity
	// closest to the portfolio's vesting date ("Tillgänglig från" is sometimes off).
	delivered := map[Date]Date{}
	var forms []Event
	for _, e := range events {
		if e.Kind == "Förm" {
			forms = append(forms, e)
		}
	}
	used := map[int]bool{}
	absDays := func(a, b Date) int {
		n := a.DaysUntil(b)
		if n < 0 {
			return -n
		}
		return n
	}
	for _, r := range portfolio {
		if r["Typ av instrument"] != "begränsade aktieenheter" || r["Tillgänglig från"] == "" {
			continue
		}
		alloc, vest := excelDate(r["Allokeringsdatum"]), excelDate(r["Tillgänglig från"])
		q := mustFloat(r["Allokerad kvantitet"])
		best := -1
		for i, e := range forms {
			if used[i] || math.Abs(e.Qty-q) >= 1e-4 || absDays(e.Date, vest) > deliveryWindowDays {
				continue
			}
			if best < 0 || absDays(e.Date, vest) < absDays(forms[best].Date, vest) {
				best = i
			}
		}
		if best >= 0 {
			used[best] = true
			delivered[alloc] = forms[best].Date
		}
	}
	// Purchased shares are free when the matching shares are delivered, or at the latest
	// after the three-year holding period.
	for i, l := range restrictedLots {
		rel := l.Vest
		if dd, ok := delivered[l.Alloc]; ok && dd.Before(rel) {
			rel = dd
		}
		if hy := addYears(l.Alloc, holdingYears); hy.Before(rel) {
			rel = hy
		}
		restrictedLots[i].Vest = rel
	}

	// Same-day benefit rows with different prices.
	byDate := map[Date]map[float64]bool{}
	for _, e := range events {
		if e.Kind == "Förm" && e.Qty != 0 {
			if byDate[e.Date] == nil {
				byDate[e.Date] = map[float64]bool{}
			}
			byDate[e.Date][math.Round(e.Cost/e.Qty*100)/100] = true
		}
	}
	var dates []Date
	for d := range byDate {
		dates = append(dates, d)
	}
	sort.Slice(dates, func(i, j int) bool { return dates[i].Before(dates[j]) })
	for _, d := range dates {
		if len(byDate[d]) > 1 {
			var ps []string
			var fs []float64
			for p := range byDate[d] {
				fs = append(fs, p)
			}
			sort.Float64s(fs)
			for _, p := range fs {
				ps = append(ps, g(p))
			}
			warn("%s: flera Förm-rader med olika pris [%s] samma dag - "+
				"stäm av förmånsvärdet mot kontrolluppgiften.", d, strings.Join(ps, ", "))
		}
	}

	ep := &EquatePlus{Events: events}
	soldInEP := false
	for _, t := range transactions {
		switch {
		case t["Ordertyp"] == "Överför" && t["Status"] == "Utförda" && strings.HasPrefix(t["Produkttyp"], "aktie"):
			ep.Transfers = append(ep.Transfers, Transfer{excelDate(t["Datum"]), mustFloat(t["Antal"]), t["Orderreferens"]})
		case t["Ordertyp"] == "Utdelning" && t["Produkttyp"] == "aktier" && t["Status"] == "Utförda":
			ep.DividendHoldings = append(ep.DividendHoldings, Holding{excelDate(t["Datum"]), mustFloat(t["Antal"])})
		case strings.HasPrefix(t["Ordertyp"], "Sälj") && t["Status"] == "Utförda":
			soldInEP = true
		}
	}
	if soldInEP {
		warn("Genomförda försäljningar i EquatePlus finns men hanteras inte av verktyget.")
	}
	for _, r := range portfolio {
		out := numOr0(r["Utestående antal"])
		switch r["Typ av instrument"] {
		case "begränsade aktieenheter":
			if out > 0 {
				ep.Pending = append(ep.Pending, Pending{excelDate(r["Allokeringsdatum"]), excelDate(r["Tillgänglig från"]), out})
			}
		case "aktier":
			ep.Current += out
			if r["Typ av insättning"] == "Köp" {
				ep.RestrictedNow += out
			}
		}
	}
	for _, p := range ep.Pending {
		ep.MatchNow += p.Qty
	}

	// Unvested matching units, in units valid on the allocation date (EquatePlus restated
	// the same allocations as the purchases).
	for _, r := range portfolio {
		if r["Typ av instrument"] == "begränsade aktieenheter" && r["Tillgänglig från"] != "" {
			alloc := excelDate(r["Allokeringsdatum"])
			q := mustFloat(r["Allokerad kvantitet"])
			if restatedOn[alloc] {
				q /= splitRatio
			}
			vest, ok := delivered[alloc]
			if !ok {
				vest = excelDate(r["Tillgänglig från"])
			}
			ep.MatchLots = append(ep.MatchLots, Lot{alloc, vest, q})
		}
	}
	ep.RestrictedLots = restrictedLots

	for _, r := range rawPF {
		if strings.TrimSpace(r["A"]) == "Per den" && r["B"] != "" {
			if f := strings.Fields(strings.ReplaceAll(r["B"], ".", "")); len(f) >= 3 {
				if m, ok := svMonths[strings.ToLower(f[1])]; ok {
					d, e1 := strconv.Atoi(f[0])
					y, e2 := strconv.Atoi(f[2])
					if e1 == nil && e2 == nil {
						ep.AsOf = mkDate(y, m, d)
					}
				}
			}
			break
		}
	}
	return ep
}

// Lot is an allocation that is locked/unvested from Alloc until Vest; Qty is in units
// valid on the allocation date.
type Lot struct {
	Alloc, Vest Date
	Qty         float64
}

// lotsHeld returns the quantity of lots allocated on/before day and still not vested at
// the end of day, in units valid on day.
func lotsHeld(lots []Lot, splitDate, day Date) float64 {
	total := 0.0
	for _, l := range lots {
		if !l.Alloc.After(day) && day.Before(l.Vest) {
			q := l.Qty
			if l.Alloc.Before(splitDate) && !day.Before(splitDate) {
				q *= splitRatio
			}
			total += q
		}
	}
	return total
}

// equatePlusBalance returns shares held in EquatePlus just before `before`
// (in units valid on that date).
func equatePlusBalance(ep *EquatePlus, splitDate, before Date) float64 {
	type item struct {
		d     Date
		ord   int
		q     float64
		split bool
	}
	var items []item
	for _, e := range ep.Events {
		items = append(items, item{e.Date, e.Order(), e.Qty, false})
	}
	for _, t := range ep.Transfers {
		items = append(items, item{t.Date, 5, -t.Qty, false})
	}
	items = append(items, item{splitDate, order["S"], 0, true})
	sort.SliceStable(items, func(i, j int) bool {
		if c := items[i].d.Cmp(items[j].d); c != 0 {
			return c < 0
		}
		return items[i].ord < items[j].ord
	})
	bal := 0.0
	for _, it := range items {
		if !it.d.Before(before) {
			break
		}
		if it.split {
			bal *= splitRatio
		} else {
			bal += it.q
		}
	}
	return bal
}

// ---------------------------------------------------------------------------- Nordnet

func loadNordnet(paths []string) *Nordnet {
	nn := &Nordnet{Files: paths}
	seen := map[string]bool{}
	var rows []map[string]string
	for _, p := range paths {
		for _, r := range readNordnet(p) {
			if !strings.HasPrefix(strings.ToUpper(r["Värdepapper"]), securityPrefix) {
				continue
			}
			if r["Makuleringsdatum"] != "" {
				continue
			}
			key := r["Id"]
			if key == "" {
				var parts []string
				for k, v := range r {
					parts = append(parts, k+"="+v)
				}
				sort.Strings(parts)
				key = strings.Join(parts, "|")
			}
			if seen[key] {
				continue
			}
			seen[key] = true
			rows = append(rows, r)
		}
	}
	if len(rows) == 0 {
		return nn
	}

	tdate := func(r map[string]string) Date {
		s := r["Affärsdag"]
		if s == "" {
			s = r["Bokföringsdag"]
		}
		d, err := parseISO(s)
		if err != nil {
			fatal("Nordnet: ogiltigt datum %q", s)
		}
		return d
	}
	rid := func(r map[string]string) int64 {
		n, _ := strconv.ParseInt(r["Id"], 10, 64)
		return n
	}
	sort.SliceStable(rows, func(i, j int) bool {
		if c := tdate(rows[i]).Cmp(tdate(rows[j])); c != 0 {
			return c < 0
		}
		return rid(rows[i]) < rid(rows[j])
	})

	// delta returns the holding change of a row; ok=false for split rows.
	delta := func(r map[string]string) (float64, bool) {
		t := strings.ToUpper(r["Transaktionstyp"])
		q := numOr0(r["Antal"])
		switch {
		case strings.Contains(t, "SPLIT"):
			return 0, false
		case t == "KÖPT" || t == "INLÄGG VP" || t == "OMVANDLING INLÄGG VP":
			return q, true
		case t == "SÅLT" || t == "UTTAG VP" || t == "OMVANDLING UTTAG VP":
			return -q, true
		}
		return 0, true
	}

	first := rows[0]
	nn.StartDate = tdate(first)
	if strings.Contains(strings.ToUpper(first["Transaktionstyp"]), "SPLIT INLÄGG") {
		nn.Opening = numOr0(first["Antal"]) / splitRatio
	} else if after, ok := num(first["Totalt antal"]); ok {
		if d0, ok := delta(first); ok {
			nn.Opening = after - d0
		}
	}

	holdings := nn.Opening
	eod := map[Date]float64{} // holdings at end of each trade date
	for _, r := range rows {
		t := strings.ToUpper(r["Transaktionstyp"])
		d := tdate(r)
		q := numOr0(r["Antal"])
		text := r["Transaktionstext"]
		switch t {
		case "SPLIT INLÄGG VP":
			nn.SplitDate = d
			holdings = q
		case "SPLIT UTTAG VP", "UTDELNING", "PRELSKATT UTD":
		case "SÅLT":
			amount := numOr0(r["Belopp"])
			fee := numOr0(r["Total Avgift"])
			price := numOr0(r["Kurs"])
			nn.Events = append(nn.Events, Event{Date: d, Kind: "Förs", Qty: q, Proceeds: amount,
				Note: fmt.Sprintf("Nordnet sålt %s st à %s kr, avgift %s kr", g(q), g(price), g(fee)),
				Ref:  r["Notanummer"]})
			holdings -= q
		case "KÖPT":
			amount := -numOr0(r["Belopp"])
			nn.Events = append(nn.Events, Event{Date: d, Kind: "Köp", Qty: q, Cost: amount,
				Note: fmt.Sprintf("Nordnet köpt %s st", g(q)), Ref: r["Notanummer"]})
			holdings += q
		case "OMVANDLING INLÄGG VP":
			cost, ok := num(r["Inköpsvärde"])
			if !ok {
				warn("Nordnet %s: OMVANDLING %s st utan inköpsvärde - omkostnadsbelopp 0 "+
					"används, ange rätt belopp manuellt.", d, g(q))
				cost = 0
			} else {
				warn("Nordnet %s: OMVANDLING %s st ('%s') medtagen som Övr med "+
					"inköpsvärde %s kr enligt Nordnet - kontrollera ursprunget.", d, g(q), text, g(cost))
			}
			nn.Events = append(nn.Events, Event{Date: d, Kind: "Övr", Qty: q, Cost: cost,
				Note: fmt.Sprintf("Nordnet omvandling %s st, inköpsvärde %s kr", g(q), g(cost))})
			holdings += q
		case "OMVANDLING UTTAG VP":
			warn("Nordnet %s: OMVANDLING UTTAG %s st ('%s') - hanteras inte, kontrollera manuellt.", d, g(q), text)
			holdings -= q
		case "INLÄGG VP":
			sq := strings.ReplaceAll(strings.ToLower(text), " ", "")
			if strings.Contains(sq, "transfer") || strings.Contains(sq, "computershare") || strings.HasSuffix(sq, "cs") {
				nn.TransfersIn = append(nn.TransfersIn, Transfer{d, q, text})
			} else {
				warn("Nordnet %s: INLÄGG VP %s st ('%s') från okänd källa - "+
					"ej med i beräkningen, lägg till manuellt om det är ett förvärv.", d, g(q), text)
			}
			holdings += q
		case "UTTAG VP":
			warn("Nordnet %s: UTTAG VP %s st ('%s') - ej försäljning, "+
				"kontrollera vart aktierna flyttades.", d, g(q), text)
			holdings -= q
		default:
			warn("Nordnet %s: okänd transaktionstyp '%s' ignorerad.", d, r["Transaktionstyp"])
		}
		if tot, ok := num(r["Totalt antal"]); ok &&
			t != "UTDELNING" && t != "PRELSKATT UTD" && t != "SPLIT UTTAG VP" &&
			math.Abs(tot-holdings) > tol {
			warn("Nordnet %s: beräknat innehav %s men 'Totalt antal' %s.", d, g(holdings), g(tot))
			holdings = tot
		}
		eod[d] = holdings
	}
	nn.Eod = eod
	nn.Closing = holdings
	nn.ClosingDate = tdate(rows[len(rows)-1])
	return nn
}

// matchTransfers pairs EquatePlus transfers out with Nordnet transfers in (same qty,
// within transferWindowDays, earliest first). It returns the number of shares that
// left EquatePlus before the Nordnet export starts but arrived in Nordnet after.
func matchTransfers(ep *EquatePlus, nn *Nordnet) float64 {
	if nn.StartDate.IsZero() {
		return 0
	}
	nn.MatchedOut, nn.MatchedIn = map[string]bool{}, map[int]bool{}
	var unmatched []int // indices into nn.TransfersIn, by date
	for i := range nn.TransfersIn {
		unmatched = append(unmatched, i)
	}
	sort.SliceStable(unmatched, func(i, j int) bool {
		return nn.TransfersIn[unmatched[i]].Date.Before(nn.TransfersIn[unmatched[j]].Date)
	})
	out := append([]Transfer(nil), ep.Transfers...)
	sort.SliceStable(out, func(i, j int) bool { return out[i].Date.Before(out[j].Date) })
	inTransit := 0.0
	for _, t := range out {
		if t.Date.Before(nn.StartDate.AddDays(-transferWindowDays)) {
			continue
		}
		hit := -1
		for k, idx := range unmatched {
			x := nn.TransfersIn[idx]
			days := t.Date.DaysUntil(x.Date)
			if math.Abs(x.Qty-t.Qty) < tol && days >= 0 && days <= transferWindowDays {
				hit = k
				break
			}
		}
		if hit >= 0 {
			nn.MatchedIn[unmatched[hit]] = true
			nn.MatchedOut[t.Ref] = true
			unmatched = append(unmatched[:hit], unmatched[hit+1:]...)
			if t.Date.Before(nn.StartDate) {
				inTransit += t.Qty
			}
		} else if !t.Date.Before(nn.StartDate) {
			warn("EquatePlus-överföring %s %s st (%s) hittades inte i Nordnet.", t.Date, g(t.Qty), t.Ref)
		}
	}
	for _, idx := range unmatched {
		x := nn.TransfersIn[idx]
		warn("Nordnet-inlägg %s %s st ('%s') saknar motsvarande EquatePlus-överföring.", x.Date, g(x.Qty), x.Ref)
	}
	return inTransit
}

// ---------------------------------------------------------------------------- calculation

type Row struct {
	Name      string
	Date      Date
	Kind      string
	Delta     float64
	TotalQty  float64
	Change    float64
	TotalCost float64
	Avg       float64
	Note      string
}

type Sale struct {
	Date     Date
	Qty      float64
	Name     string
	Proceeds float64
	Cost     float64
	Result   float64
	Ref      string
	// LedgerAfter is the SKV 2198 total after the sale.
	LedgerAfter float64
}

func compute(events []Event, name string) ([]Row, []Sale) {
	ev := append([]Event(nil), events...)
	sortEvents(ev)
	var rows []Row
	var sales []Sale
	totalQty, totalCost := 0.0, 0.0
	for _, e := range ev {
		var delta, change float64
		switch e.Kind {
		case "S":
			delta = totalQty * (splitRatio - 1)
		case "Förs", "Adj":
			if e.Qty > totalQty+tol {
				warn("%s: försäljning av %s st men endast %s st i beräkningen - "+
					"historiken är ofullständig.", e.Date, g(e.Qty), g(totalQty))
			}
			avgBefore := 0.0
			if totalQty != 0 {
				avgBefore = totalCost / totalQty
			}
			delta, change = -e.Qty, -e.Qty*avgBefore
			if e.Kind == "Förs" {
				sales = append(sales, Sale{Date: e.Date, Qty: e.Qty, Name: name, Proceeds: e.Proceeds,
					Cost: -change, Result: e.Proceeds + change, Ref: e.Ref})
			}
		default:
			delta, change = e.Qty, e.Cost
		}
		totalQty += delta
		totalCost += change
		if e.Kind == "Förs" {
			sales[len(sales)-1].LedgerAfter = totalQty
		}
		if math.Abs(totalQty) < tol {
			totalQty, totalCost = 0, 0
		}
		avg := 0.0
		if totalQty != 0 {
			avg = totalCost / totalQty
		}
		kind := e.Kind
		if kind == "Adj" {
			kind = "Övr"
		}
		rows = append(rows, Row{name, e.Date, kind, delta, totalQty, change, totalCost, avg, e.Note})
	}
	return rows, sales
}

// ledgerTotalBefore returns total shares in the ledger before (before, ord).
func ledgerTotalBefore(events []Event, before Date, ord int) float64 {
	ev := append([]Event(nil), events...)
	sortEvents(ev)
	tot := 0.0
	for _, e := range ev {
		if c := e.Date.Cmp(before); c > 0 || (c == 0 && e.Order() >= ord) {
			break
		}
		switch e.Kind {
		case "S":
			tot *= splitRatio
		case "Förs", "Adj":
			tot -= e.Qty
		default:
			tot += e.Qty
		}
	}
	return tot
}

// ---------------------------------------------------------------------------- holdings log

// Same-day order: everything else before sales (as in the SKV 2198 table), transfers
// out of EquatePlus before arrivals in Nordnet.
var logOrder = map[string]int{"Adj": -1, "S": 0, "Grant": 1, "Vest": 1, "Köp": 1, "Förm": 2, "Övr": 2,
	"Out": 4, "In": 5, "Förs": 9}

var logLabel = map[string]string{"Adj": "Justering", "S": "Split", "Grant": "Tilldelning matchning",
	"Vest": "Intjänande", "Köp": "Köp", "Förm": "Förm", "Övr": "Övr",
	"Out": "Överföring ut", "In": "Överföring in", "Förs": "Försäljning"}

var logHeader = []string{"Inkomstår", "Datum", "Händelse", "Källa", "Antal", "Beteckning",
	"Försäljningspris", "Omkostnadsbelopp", "Vinst", "Förlust", "Nota", "Kommentar",
	"EQP Tillgänglig för handel", "EQP Tillgänglig med begränsningar",
	"EQP Företagsmatchning (ej intjänad, ingår ej i totalt)", "Under överföring",
	"Kvar Nordnet", "Kvar totalt"}

type logItem struct {
	date       Date
	order, seq int
	kind       string
	qty        float64
	source     string
	text       string
	locked     float64 // Köp: shares locked until the matching shares vest
	matched    bool    // Out/In: transfer matched between EquatePlus and Nordnet
	rlots      []Lot   // Vest: purchase lots released
	mlots      []Lot   // Vest: matching lots vested
}

// LogRow is one row of the holdings log.
type LogRow struct {
	Date                                      Date
	Kind, Source, Name, Comment               string
	Change                                    float64
	Sale                                      *Sale
	Free, Restricted, Match, Transit, Nordnet float64
	Total                                     float64
}

func (r LogRow) cells(nf numFmt) []string {
	c := []string{r.Date.Year4(), r.Date.String(), logLabel[r.Kind], r.Source, nf.f(r.Change, 5), r.Name,
		"", "", "", "", "", r.Comment}
	if s := r.Sale; s != nil {
		c[6], c[7] = nf.f(s.Proceeds, 0), nf.f(s.Cost, 0)
		c[8], c[9] = nf.f(math.Max(s.Result, 0), 0), nf.f(math.Max(-s.Result, 0), 0)
		c[10] = s.Ref
	}
	return append(c, nf.f(r.Free, 5), nf.f(r.Restricted, 5), nf.f(r.Match, 5), nf.f(r.Transit, 5),
		nf.f(r.Nordnet, 5), nf.f(r.Total, 5))
}

// buildLog returns one row per event that changes holdings anywhere, with all holdings
// after it.
func buildLog(ep *EquatePlus, nn *Nordnet, ledger []Event, sales []Sale, splitDate Date, name string) []LogRow {
	cutoff := ep.AsOf
	if cutoff.IsZero() {
		cutoff = nn.ClosingDate
		for _, e := range ledger {
			if e.Date.After(cutoff) {
				cutoff = e.Date
			}
		}
	}
	var items []logItem
	add := func(it logItem) {
		it.order, it.seq = logOrder[it.kind], len(items)
		items = append(items, it)
	}

	locked := map[Date]float64{}
	for _, l := range ep.RestrictedLots {
		locked[l.Alloc] = l.Qty
	}
	for _, e := range ep.Events {
		it := logItem{date: e.Date, kind: e.Kind, qty: e.Qty, source: "EquatePlus", text: e.Note}
		if e.Kind == "Köp" {
			it.locked = locked[e.Date]
		}
		add(it)
	}
	type vestLots struct{ r, m []Lot }
	vests := map[Date]*vestLots{}
	var vestDates []Date
	vestOf := func(d Date) *vestLots {
		if vests[d] == nil {
			vests[d] = &vestLots{}
			vestDates = append(vestDates, d)
		}
		return vests[d]
	}
	for _, l := range ep.MatchLots {
		if !l.Alloc.After(cutoff) {
			add(logItem{date: l.Alloc, kind: "Grant", qty: l.Qty, source: "EquatePlus",
				text: fmt.Sprintf("Företagsmatchning tilldelad, intjänas %s", l.Vest)})
		}
		if !l.Vest.After(cutoff) {
			v := vestOf(l.Vest)
			v.m = append(v.m, l)
		}
	}
	for _, l := range ep.RestrictedLots {
		if !l.Vest.After(cutoff) {
			v := vestOf(l.Vest)
			v.r = append(v.r, l)
		}
	}
	for _, d := range vestDates {
		add(logItem{date: d, kind: "Vest", source: "EquatePlus", rlots: vests[d].r, mlots: vests[d].m})
	}
	for _, t := range ep.Transfers {
		add(logItem{date: t.Date, kind: "Out", qty: t.Qty, source: "EquatePlus",
			text: fmt.Sprintf("Till Nordnet (%s)", t.Ref), matched: nn.MatchedOut[t.Ref]})
	}
	for i, t := range nn.TransfersIn {
		add(logItem{date: t.Date, kind: "In", qty: t.Qty, source: "Nordnet",
			text: fmt.Sprintf("Från EquatePlus (%s)", t.Ref), matched: nn.MatchedIn[i]})
	}
	for _, e := range nn.Events {
		add(logItem{date: e.Date, kind: e.Kind, qty: e.Qty, source: "Nordnet", text: e.Note})
	}
	for _, e := range ledger {
		if e.Kind == "S" || e.Kind == "Adj" {
			add(logItem{date: e.Date, kind: e.Kind, qty: e.Qty, text: e.Note})
		}
	}
	sort.SliceStable(items, func(i, j int) bool {
		a, b := items[i], items[j]
		if c := a.date.Cmp(b.date); c != 0 {
			return c < 0
		}
		if a.order != b.order {
			return a.order < b.order
		}
		return a.seq < b.seq
	})

	units := func(l Lot, day Date) float64 {
		if l.Alloc.Before(splitDate) && !day.Before(splitDate) {
			return l.Qty * splitRatio
		}
		return l.Qty
	}

	var epTot, restr, match, transit, nord float64
	saleIdx := 0
	nordEod := map[Date]float64{}
	var rows []LogRow
	for _, it := range items {
		change, text := it.qty, it.text
		var sale *Sale
		before := epTot + transit + nord
		switch {
		case it.kind == "S":
			epTot, restr, match, transit, nord = epTot*splitRatio, restr*splitRatio, match*splitRatio,
				transit*splitRatio, nord*splitRatio
			change = epTot + transit + nord - before
		case it.kind == "Grant":
			match += it.qty
		case it.kind == "Vest":
			rq, mq := 0.0, 0.0
			for _, l := range it.rlots {
				rq += units(l, it.date)
			}
			for _, l := range it.mlots {
				mq += units(l, it.date)
			}
			restr -= rq
			match -= mq
			change = rq
			text = fmt.Sprintf("Begränsning upphör för %s st köpta aktier; %s st företagsmatchning "+
				"intjänad (levereras som Förm)", fmtSv(rq, 5), fmtSv(mq, 5))
		case it.source == "EquatePlus" && (it.kind == "Köp" || it.kind == "Förm" || it.kind == "Övr"):
			epTot += it.qty
			restr += it.locked
		case it.kind == "Out":
			epTot -= it.qty
			transit += it.qty
			change = -it.qty
			if !it.matched {
				text += " - ej hittad i Nordnet"
			}
		case it.kind == "In":
			nord += it.qty
			if it.matched {
				transit -= it.qty
			} else {
				text += " - ingen motsvarande EquatePlus-överföring"
			}
		case it.source == "Nordnet" && (it.kind == "Köp" || it.kind == "Övr"):
			nord += it.qty
		case it.kind == "Förs":
			nord -= it.qty
			change = -it.qty
			if saleIdx < len(sales) {
				sale = &sales[saleIdx]
				saleIdx++
			}
		case it.kind == "Adj":
			transit -= it.qty
			change = -it.qty
		}
		free := epTot - restr
		total := epTot + transit + nord
		if free < -0.01 {
			warn("%s: fler begränsade aktier (%.5f) än innehav i EquatePlus (%.5f).", it.date, restr, epTot)
		}
		if sale != nil && math.Abs(total-sale.LedgerAfter) > 0.01 {
			warn("%s: innehav efter försäljning %.5f st stämmer inte med SKV 2198-tabellens %.5f st.",
				it.date, total, sale.LedgerAfter)
		}
		nordEod[it.date] = nord
		rows = append(rows, LogRow{Date: it.date, Kind: it.kind, Source: it.source, Name: name,
			Comment: text, Change: change, Sale: sale, Free: free, Restricted: restr, Match: match,
			Transit: transit, Nordnet: nord, Total: total})
	}

	// Cross-checks of the simulated holdings.
	var eodDates []Date
	for d := range nn.Eod {
		eodDates = append(eodDates, d)
	}
	sort.Slice(eodDates, func(i, j int) bool { return eodDates[i].Before(eodDates[j]) })
	for _, d := range eodDates {
		if v, ok := nordEod[d]; ok && math.Abs(v-nn.Eod[d]) > tol {
			warn("Innehavslogg %s: Nordnet %s st, enligt Nordnet %s st.", d, g(v), g(nn.Eod[d]))
		}
	}
	if !ep.AsOf.IsZero() {
		for _, c := range []struct {
			label      string
			calc, real float64
		}{
			{"EquatePlus totalt", epTot, ep.Current},
			{"begränsade", restr, ep.RestrictedNow},
			{"ej intjänad företagsmatchning", match, ep.MatchNow},
		} {
			if math.Abs(c.calc-c.real) > 0.01 {
				warn("Innehavslogg %s: %s %.5f st, enligt portföljen %.5f st.", ep.AsOf, c.label, c.calc, c.real)
			}
		}
	}
	if !nn.StartDate.IsZero() && math.Abs(nord-nn.Closing) > tol {
		warn("Innehavslogg: Nordnet slutsaldo %s st, enligt Nordnet %s st.", g(nord), g(nn.Closing))
	}
	return rows
}

// ---------------------------------------------------------------------------- output

type numFmt struct {
	swedish bool
}

func (n numFmt) f(v float64, dec int) string {
	if n.swedish {
		return strings.ReplaceAll(fmtSv(v, dec), " ", "")
	}
	return strconv.FormatFloat(v, 'f', dec, 64)
}

func writeCSV(path string, header []string, rows [][]string, swedish bool) {
	f, err := os.Create(path)
	if err != nil {
		fatal("%s: %v", path, err)
	}
	defer f.Close()
	f.WriteString("\ufeff")
	w := csv.NewWriter(f)
	w.Comma = ','
	if swedish {
		w.Comma = ';'
	}
	w.UseCRLF = true
	w.Write(header)
	w.WriteAll(rows)
	if err := w.Error(); err != nil {
		fatal("%s: %v", path, err)
	}
}

func printTable(headers []string, data [][]string, left map[int]bool) {
	widths := make([]int, len(headers))
	for i, h := range headers {
		widths[i] = utf8.RuneCountInString(h)
		for _, d := range data {
			if n := utf8.RuneCountInString(d[i]); n > widths[i] {
				widths[i] = n
			}
		}
	}
	line := func(cells []string) string {
		parts := make([]string, len(cells))
		for i, c := range cells {
			pad := strings.Repeat(" ", widths[i]-utf8.RuneCountInString(c))
			if left[i] {
				parts[i] = c + pad
			} else {
				parts[i] = pad + c
			}
		}
		return strings.Join(parts, "  ")
	}
	fmt.Println(line(headers))
	dashes := make([]string, len(widths))
	for i, w := range widths {
		dashes[i] = strings.Repeat("-", w)
	}
	fmt.Println(strings.Join(dashes, "  "))
	for _, d := range data {
		fmt.Println(line(d))
	}
}

type multiFlag []string

func (m *multiFlag) String() string     { return strings.Join(*m, ",") }
func (m *multiFlag) Set(s string) error { *m = append(*m, s); return nil }

type dateFlag struct{ d Date }

func (f *dateFlag) String() string {
	if f.d.IsZero() {
		return ""
	}
	return f.d.String()
}
func (f *dateFlag) Set(s string) error {
	d, err := parseISO(s)
	f.d = d
	return err
}

// splitArgs lets flags appear after the positional directory argument.
func splitArgs(args []string, fs *flag.FlagSet) (flagsPart, positional []string) {
	for i := 0; i < len(args); i++ {
		a := args[i]
		if a == "--" {
			positional = append(positional, args[i+1:]...)
			break
		}
		if strings.HasPrefix(a, "-") && len(a) > 1 {
			flagsPart = append(flagsPart, a)
			name := strings.TrimLeft(a, "-")
			if strings.Contains(name, "=") {
				continue
			}
			if f := fs.Lookup(name); f != nil {
				if bf, ok := f.Value.(interface{ IsBoolFlag() bool }); ok && bf.IsBoolFlag() {
					continue
				}
				if i+1 < len(args) {
					i++
					flagsPart = append(flagsPart, args[i])
				}
			}
			continue
		}
		positional = append(positional, a)
	}
	return
}

func main() {
	fs := flag.NewFlagSet("eqpskv", flag.ExitOnError)
	output := fs.String("o", "skv2198_saab.csv", "SKV 2198 CSV")
	k4Path := fs.String("k4", "k4_saab.csv", "K4 sales CSV")
	k4HoldPath := fs.String("k4-holdings", "k4_innehav_saab.csv", "holdings log CSV (every event, holdings after each)")
	year := fs.Int("year", 0, "show K4 sales only for this income year")
	name := fs.String("name", "Saab B", "Beteckning")
	var nordnetFiles multiFlag
	fs.Var(&nordnetFiles, "nordnet", "Nordnet export file (repeatable); default: Nordnet*.csv in the folder")
	var splitFlag dateFlag
	fs.Var(&splitFlag, "split-date", fmt.Sprintf("date of the 4:1 split, YYYY-MM-DD (default: from Nordnet, else %s)", fallbackSplitDate))
	excludeSPA := fs.Bool("exclude-share-plan-account", false, "leave out the 'Share Plan Account' allocation")
	noAdjust := fs.Bool("no-adjust", false, "do not insert an adjustment row for shares unaccounted for")
	decimalPoint := fs.Bool("decimal-point", false, "write CSV with '.' decimals and ',' separator")
	fs.Usage = func() {
		fmt.Fprintf(fs.Output(), "Usage: eqpskv [flags] [directory]\n\n"+
			"Builds an SKV 2198 table and K4 summary for Saab B from EquatePlus and Nordnet exports.\n\n")
		fs.PrintDefaults()
	}
	flagArgs, pos := splitArgs(os.Args[1:], fs)
	fs.Parse(flagArgs)
	dir := "."
	if len(pos) > 0 {
		dir = pos[0]
	}

	nnPaths := []string(nordnetFiles)
	if len(nnPaths) == 0 {
		nnPaths = find(dir, "Nordnet*.csv")
	}
	nn := &Nordnet{}
	if len(nnPaths) > 0 {
		nn = loadNordnet(nnPaths)
	}
	splitDate := splitFlag.d
	if splitDate.IsZero() {
		splitDate = nn.SplitDate
	}
	if splitDate.IsZero() {
		splitDate = fallbackSplitDate
		warn("Splitdatum okänt, antar %s (ange -split-date).", splitDate)
	}

	ep := loadEquatePlus(dir, splitDate, *excludeSPA)
	events := append(append([]Event(nil), ep.Events...), nn.Events...)
	if len(ep.Events) > 0 {
		minD := ep.Events[0].Date
		for _, e := range ep.Events {
			if e.Date.Before(minD) {
				minD = e.Date
			}
		}
		if minD.Before(splitDate) {
			events = append(events, Event{Date: splitDate, Kind: "S", Note: fmt.Sprintf("Split %d:1", splitRatio)})
		}
	}

	// Check the restricted / unvested split against the portfolio per its "Per den" date.
	if !ep.AsOf.IsZero() {
		if r := lotsHeld(ep.RestrictedLots, splitDate, ep.AsOf); math.Abs(r-ep.RestrictedNow) > 0.01 {
			warn("Begränsade aktier %s: beräknat %.5f st, enligt portföljen %.5f st.", ep.AsOf, r, ep.RestrictedNow)
		}
		if m := lotsHeld(ep.MatchLots, splitDate, ep.AsOf); math.Abs(m-ep.MatchNow) > 0.01 {
			warn("Ej intjänad företagsmatchning %s: beräknat %.5f st, enligt portföljen %.5f st.", ep.AsOf, m, ep.MatchNow)
		}
	}

	// Check EquatePlus holdings against the dividend share counts.
	for _, h := range ep.DividendHoldings {
		if b := equatePlusBalance(ep, splitDate, h.Date); math.Abs(b-h.Qty) > 0.01 {
			warn("EquatePlus innehav %s: beräknat %.5f st, enligt utdelning %.5f st.", h.Date, b, h.Qty)
		}
	}

	// Reconcile with the Nordnet opening balance; adjust for shares that left earlier.
	if !nn.StartDate.IsZero() {
		inTransit := matchTransfers(ep, nn)
		start := nn.StartDate
		ledger := ledgerTotalBefore(events, start, order["Adj"])
		outside := ledger - equatePlusBalance(ep, splitDate, start) - inTransit
		adjQty := outside - nn.Opening
		if adjQty > 0.01 && !*noAdjust {
			events = append(events, Event{Date: start, Kind: "Adj", Qty: adjQty,
				Note: "Justering: aktier avyttrade/flyttade före Nordnet-exportens början (okänd tidpunkt)"})
			warn("%.5f st (räknat i antal före split om datumet ligger före split) som förts ut från "+
				"EquatePlus finns inte i Nordnet vid exportens början %s. De tas bort som 'Övr' till "+
				"genomsnittligt omkostnadsbelopp på det datumet. Ta med en Nordnet-export som täcker "+
				"hela historiken (även gamla ISIN SE0000112385) eller redovisa tidigare försäljningar "+
				"för ett exakt resultat.", adjQty, start)
		} else if adjQty < -0.01 {
			warn("Nordnet hade %.5f st fler aktier vid exportens början än EquatePlus förklarar - "+
				"förvärv från annan källa saknas i beräkningen.", -adjQty)
		}
	}

	rows, sales := compute(events, *name)
	if len(rows) == 0 {
		fatal("inga transaktioner hittades")
	}

	fmt.Printf("SKV 2198 - Genomsnittligt omkostnadsbelopp - %s\n\n", *name)
	var data [][]string
	for _, r := range rows {
		data = append(data, []string{r.Date.String(), r.Kind, fmtSv(r.Delta, 5), fmtSv(r.TotalQty, 5),
			fmtSv(r.Change, 2), fmtSv(r.TotalCost, 2), fmtSv(r.Avg, 2), r.Note})
	}
	printTable([]string{"Datum", "Händ.", "Antal +/-", "Totalt antal", "Inköpspris/utnyttjat",
		"Totalt omk.belopp", "Genomsn. omk.belopp", "Kommentar"}, data, map[int]bool{0: true, 1: true, 7: true})

	nf := numFmt{swedish: !*decimalPoint}
	var csvRows [][]string
	for _, r := range rows {
		csvRows = append(csvRows, []string{r.Name, r.Date.String(), r.Kind, nf.f(r.Delta, 5), nf.f(r.TotalQty, 5),
			nf.f(r.Change, 2), nf.f(r.TotalCost, 2), nf.f(r.Avg, 2), r.Note})
	}
	writeCSV(*output, []string{"Beteckning", "Datum", "Händelse", "Antal +/-", "Totalt antal",
		"Inköpspris (+) / utnyttjat omk.belopp (-), kr", "Totalt outnyttjat omk.belopp, kr",
		"Genomsnittligt omk.belopp, kr", "Kommentar"}, csvRows, nf.swedish)
	fmt.Printf("\nSKV 2198-tabell sparad: %s\n", *output)

	// K4
	if len(sales) > 0 {
		var shown []Sale
		for _, s := range sales {
			if *year == 0 || s.Date.Year() == *year {
				shown = append(shown, s)
			}
		}
		var years []string
		seenY := map[string]bool{}
		for _, s := range shown {
			if y := s.Date.Year4(); !seenY[y] {
				seenY[y] = true
				years = append(years, y)
			}
		}
		sort.Strings(years)
		var k4 [][]string
		var disp [][]string
		pos := func(v float64) float64 { return math.Max(v, 0) }
		for _, y := range years {
			var tq, tp, to float64
			for _, s := range shown {
				if s.Date.Year4() != y {
					continue
				}
				k4 = append(k4, []string{y, s.Date.String(), nf.f(s.Qty, 5), s.Name, nf.f(s.Proceeds, 0),
					nf.f(s.Cost, 0), nf.f(pos(s.Result), 0), nf.f(pos(-s.Result), 0), s.Ref})
				disp = append(disp, []string{y, s.Date.String(), fmtSv(s.Qty, 0), fmtSv(s.Proceeds, 0),
					fmtSv(s.Cost, 0), fmtSv(pos(s.Result), 0), fmtSv(pos(-s.Result), 0)})
				tq += s.Qty
				tp += s.Proceeds
				to += s.Cost
			}
			tr := tp - to
			k4 = append(k4, []string{y, "SUMMA", nf.f(tq, 5), *name, nf.f(tp, 0), nf.f(to, 0),
				nf.f(pos(tr), 0), nf.f(pos(-tr), 0), ""})
			disp = append(disp, []string{y, "SUMMA K4", fmtSv(tq, 0), fmtSv(tp, 0), fmtSv(to, 0),
				fmtSv(pos(tr), 0), fmtSv(pos(-tr), 0)})
			disp = append(disp, make([]string, 7))
		}
		if len(disp) > 0 {
			disp = disp[:len(disp)-1]
		}
		fmt.Println("\nK4 avsnitt A - försäljningar")
		printTable([]string{"År", "Datum", "Antal", "Försäljningspris", "Omkostnadsbelopp", "Vinst", "Förlust"},
			disp, map[int]bool{0: true, 1: true})
		if len(k4) > 0 {
			writeCSV(*k4Path, []string{"Inkomstår", "Datum", "Antal", "Beteckning", "Försäljningspris",
				"Omkostnadsbelopp", "Vinst", "Förlust", "Nota"}, k4, nf.swedish)
			fmt.Printf("\nK4-underlag sparat: %s  (avrunda till hela kronor på blanketten)\n", *k4Path)
		}
	}

	// Holdings log
	logRows := buildLog(ep, nn, events, sales, splitDate, *name)
	var logCSV [][]string
	for _, r := range logRows {
		if *year != 0 && r.Date.Year() != *year {
			continue
		}
		logCSV = append(logCSV, r.cells(nf))
	}
	if len(logCSV) > 0 {
		writeCSV(*k4HoldPath, logHeader, logCSV, nf.swedish)
		fmt.Printf("Innehavslogg (alla händelser, innehav efter varje) sparad: %s\n", *k4HoldPath)
	}

	// Summary
	last := rows[len(rows)-1]
	fmt.Println("\nSammanfattning")
	fmt.Printf("  Totalt innehav enligt beräkningen:  %s st, genomsnitt %s kr/st\n",
		fmtSv(last.TotalQty, 5), fmtSv(last.Avg, 2))
	held := ep.Current
	line := fmt.Sprintf("  Faktiskt innehav: EquatePlus %s st", fmtSv(ep.Current, 5))
	if !nn.StartDate.IsZero() {
		held += nn.Closing
		line += fmt.Sprintf(" + Nordnet %s st (%s)", fmtSv(nn.Closing, 0), nn.ClosingDate)
	}
	fmt.Printf("%s = %s st\n", line, fmtSv(held, 5))
	if math.Abs(held-last.TotalQty) > 0.01 {
		warn("Beräknat innehav %.5f st skiljer sig från faktiskt %.5f st.", last.TotalQty, held)
	}
	if len(ep.Pending) > 0 {
		sum := 0.0
		minV, maxV := ep.Pending[0].Vest, ep.Pending[0].Vest
		for _, p := range ep.Pending {
			sum += p.Qty
			if p.Vest.Before(minV) {
				minV = p.Vest
			}
			if p.Vest.After(maxV) {
				maxV = p.Vest
			}
		}
		fmt.Printf("  Ej intjänade matchningsaktier (ingår ej): %s st, intjänas %s - %s\n", fmtSv(sum, 5), minV, maxV)
	}
	if nn.StartDate.IsZero() {
		fmt.Println("  Ingen Nordnet-fil: inga försäljningar med i beräkningen.")
	}
	if len(warnings) > 0 {
		fmt.Println("\nVarningar / kontrollera:")
		for _, m := range warnings {
			fmt.Println("  * " + m)
		}
	}
	fmt.Println("\nFörmånsvärdet för matchningsaktier bör stämmas av mot kontrolluppgift/lönebesked.")
}
