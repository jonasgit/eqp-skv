#!/usr/bin/env python3
"""
eqp_skv.py - Build an SKV 2198 table ("Genomsnittligt omkostnadsbelopp") and K4 sales
summary for Saab B shares acquired through EquatePlus (Saab Share Matching Plan) and
sold through Nordnet.

Uses only the Python standard library (no openpyxl/pandas needed).

Input files (looked up in the given directory):
  PortfolioDetails*.xlsx              - all allocations: purchases, vested matching shares,
                                        free shares (primary source of acquisitions)
  CompletedTransactions*.xlsx         - deliveries (Leverans), transfers out (Överför) and
                                        dividends (used as holding checkpoints)
  Sammanfattning av inköpscykel*.csv  - monthly purchases (detects allocations that
                                        EquatePlus restated after the 2024 split)
  Nordnet*.csv                        - Nordnet "transaktioner och notor" exports (one or
                                        more, UTF-16, tab separated); sales and purchases of
                                        SAAB B, old and new ISIN

Outputs:
  skv2198_saab.csv  - SKV 2198 columns 1-8 (+ comment), chronological
  k4_saab.csv       - one row per sale and per-year totals for K4 section A
  k4_innehav_saab.csv - holdings log: one row per event (purchases, deliveries, matching
                      grants and vestings, transfers, sales, split) with K4 values on sale
                      rows and holdings after each event in EquatePlus (free for trading /
                      with restrictions / unvested company match), under transfer, in
                      Nordnet and in total (units valid on the date; unvested matching is
                      not part of the total)

Tax treatment used:
  * Purchased Shares (EquatePlus) and KÖPT (Nordnet) -> "Köp", cost incl. fees.
  * Matching / Free shares -> "Förm" (förmånsbeskattat belopp): market value at delivery.
  * Saab 4:1 split -> "S": quantity x4, total cost unchanged.
  * SÅLT (Nordnet) -> "Förs": utnyttjat omkostnadsbelopp = antal x genomsnitt.
  * Transfers EquatePlus -> Nordnet are not disposals and do not appear in the table.
  * Unvested matching shares (begränsade aktieenheter) are not yet owned and are excluded.
  * If the Nordnet export does not reach back far enough to explain where all shares
    transferred out of EquatePlus went, an adjustment row ("Övr") removes the unexplained
    shares at the average cost at the start of the export, with a warning.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import glob
import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

SV_MONTHS = {
    "jan": 1, "feb": 2, "mars": 3, "mar": 3, "apr": 4, "maj": 5, "juni": 6, "jun": 6,
    "juli": 7, "jul": 7, "aug": 8, "sep": 9, "sept": 9, "okt": 10, "nov": 11, "dec": 12,
}

# Saab AB 4:1 share split (spring 2024). Taken from the Nordnet "SPLIT INLÄGG VP" row
# when present; otherwise this fallback (override with --split-date).
FALLBACK_SPLIT_DATE = dt.date(2024, 4, 19)
SPLIT_RATIO = 4
SECURITY_PREFIX = "SAAB B"
TOL = 0.001
HOLDING_YEARS = 3           # Share Matching Plan holding period for purchased shares
DELIVERY_WINDOW_DAYS = 45   # max distance between portfolio vesting date and delivery

warnings: list[str] = []


def warn(msg: str) -> None:
    warnings.append(msg)


# --------------------------------------------------------------------------- parsing


def read_xlsx(path: str) -> list[dict[str, str]]:
    """Return rows of the first sheet as {column letter: value} dicts."""
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t")))
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    rows = []
    for r in sheet.iter(f"{{{NS['m']}}}row"):
        row: dict[str, str] = {}
        for c in r.findall("m:c", NS):
            col = "".join(ch for ch in c.get("r", "") if ch.isalpha())
            v = c.find("m:v", NS)
            t = c.get("t")
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(f"{{{NS['m']}}}t"))
            else:
                val = v.text if v is not None and v.text is not None else ""
            row[col] = val
        rows.append(row)
    return rows


def table(rows: list[dict[str, str]], header_key: str) -> list[dict[str, str]]:
    """Find the header row whose column A == header_key and return rows keyed by header."""
    for i, r in enumerate(rows):
        if r.get("A", "").strip() == header_key:
            out = []
            for rr in rows[i + 1:]:
                if any(v.strip() for v in rr.values()):
                    out.append({name: rr.get(col, "") for col, name in r.items()})
            return out
    raise ValueError(f"header '{header_key}' not found")


def add_years(d: dt.date, years: int) -> dt.date:
    try:
        return d.replace(year=d.year + years)
    except ValueError:  # 29 Feb
        return d.replace(year=d.year + years, day=28)


def excel_date(v: str) -> dt.date:
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(float(v)))


def sv_date(s: str) -> dt.date:
    """Parse '16 nov. 2023' / '16 maj 2024'."""
    d, m, y = s.replace(".", "").split()
    return dt.date(int(y), SV_MONTHS[m.lower()], int(d))


def num(s: str | None) -> float | None:
    s = (s or "").strip().replace(" ", "").replace(" ", "")
    if s in ("", "-"):
        return None
    return float(s.replace(",", "."))


def read_sv_csv(path: str) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def read_nordnet(path: str) -> list[dict[str, str]]:
    raw = open(path, "rb").read()
    for enc in ("utf-16", "utf-8-sig", "cp1252"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    lines = text.splitlines()
    delim = "\t" if "\t" in lines[0] else ";"
    reader = csv.reader(lines, delimiter=delim)
    header = next(reader)
    rows = []
    for rec in reader:
        if not any(x.strip() for x in rec):
            continue
        # Header has duplicate "Valuta" columns; keep the first occurrence of each name.
        d: dict[str, str] = {}
        for name, val in zip(header, rec):
            d.setdefault(name.strip(), val.strip())
        rows.append(d)
    return rows


def find(directory: str, pattern: str) -> list[str]:
    return sorted(glob.glob(os.path.join(directory, pattern)))


# --------------------------------------------------------------------------- model

# Order of events on the same date.
ORDER = {"Adj": -1, "S": 0, "Köp": 1, "Förm": 2, "Övr": 2, "Förs": 3}


@dataclass
class Event:
    date: dt.date
    kind: str          # Köp, Förm, S, Övr, Förs, Adj (adjustment, shown as Övr)
    qty: float = 0.0   # shares in (+); for Förs/Adj the number of shares out (positive)
    cost: float = 0.0  # added omkostnadsbelopp (Köp/Förm/Övr)
    proceeds: float = 0.0  # Förs: net sale amount after fees
    note: str = ""
    ref: str = ""

    @property
    def order(self) -> int:
        return ORDER[self.kind]


@dataclass
class EquatePlus:
    events: list[Event]
    transfers: list[tuple[dt.date, float, str]]
    dividend_holdings: list[tuple[dt.date, float]]
    pending: list[tuple[dt.date, dt.date, float]]
    current: float
    # (allocation date, vesting date, qty in units valid on the allocation date)
    restricted_lots: list[tuple[dt.date, dt.date, float]] = field(default_factory=list)
    match_lots: list[tuple[dt.date, dt.date, float]] = field(default_factory=list)
    as_of: dt.date | None = None          # "Per den" date of PortfolioDetails
    restricted_now: float = 0.0           # Tillgänglig med begränsningar per as_of
    match_now: float = 0.0                # unvested matching units per as_of


def lots_held(lots, split_date: dt.date, day: dt.date) -> float:
    """Quantity of lots allocated on/before `day` and still not vested at end of `day`,
    in units valid on `day`."""
    total = 0.0
    for alloc, vest, q in lots:
        if alloc <= day < vest:
            total += q * (SPLIT_RATIO if alloc < split_date <= day else 1)
    return total


@dataclass
class Nordnet:
    events: list[Event] = field(default_factory=list)
    transfers_in: list[tuple[dt.date, float, str]] = field(default_factory=list)
    split_date: dt.date | None = None
    start_date: dt.date | None = None
    opening: float = 0.0           # holdings just before start_date (units of that date)
    closing: float = 0.0
    closing_date: dt.date | None = None
    files: list[str] = field(default_factory=list)
    eod: dict[dt.date, float] = field(default_factory=dict)  # holdings at end of trade date
    # Transfers matched between EquatePlus and Nordnet
    matched_out: set[str] = field(default_factory=set)   # EquatePlus order references
    matched_in: set[int] = field(default_factory=set)    # indices into transfers_in


# --------------------------------------------------------------------------- EquatePlus


def load_equateplus(directory: str, split_date: dt.date, exclude_spa: bool) -> EquatePlus:
    pf = find(directory, "PortfolioDetails*.xlsx")
    if not pf:
        sys.exit(f"PortfolioDetails*.xlsx not found in {directory}")
    tx = find(directory, "CompletedTransactions*.xlsx")
    cyc = find(directory, "Sammanfattning av inköpscykel*.csv")

    raw_pf = read_xlsx(pf[0])
    portfolio = table(raw_pf, "Allokeringsdatum")
    restated_on: dict[dt.date, bool] = {}
    restricted_lots: list[tuple[dt.date, dt.date, float]] = []
    transactions = table(read_xlsx(tx[0]), "Orderreferens") if tx else []
    if not tx:
        warn("CompletedTransactions*.xlsx saknas: inga kontroller mot leveranser/överföringar.")

    bought_on: dict[dt.date, float] = {}
    if cyc:
        for r in read_sv_csv(cyc[0]):
            q = num(r.get("Köpta aktier"))
            if q and r.get("Inköpsdatum"):
                d = sv_date(r["Inköpsdatum"])
                bought_on[d] = bought_on.get(d, 0.0) + q
    else:
        warn("Inköpscykel-CSV saknas: omräkning för split bedöms enbart via intjänandedatum.")

    vest_of: dict[dt.date, dt.date] = {}
    for r in portfolio:
        if r["Typ av instrument"] == "begränsade aktieenheter" and r["Tillgänglig från"]:
            vest_of[excel_date(r["Allokeringsdatum"])] = excel_date(r["Tillgänglig från"])

    deliveries: dict[dt.date, float] = {}
    for t in transactions:
        if t.get("Ordertyp") == "Leverans" and t.get("Status") == "Utförda":
            d = excel_date(t["Datum"])
            deliveries[d] = deliveries.get(d, 0.0) + float(t["Antal"])

    events: list[Event] = []
    raw_purchase_price: dict[dt.date, float] = {}

    # Purchases.
    for r in portfolio:
        if r["Typ av instrument"] != "aktier" or r["Typ av insättning"] != "Köp":
            continue
        d = excel_date(r["Allokeringsdatum"])
        qty, price = float(r["Allokerad kvantitet"]), num(r["Lösenpris/kostnadsbas"])
        if price is None:
            warn(f"{d}: köp utan pris, hoppas över ({qty} st)")
            continue
        restated = False
        if d < split_date:
            csv_q = bought_on.get(d)
            if csv_q:
                if abs(qty - SPLIT_RATIO * csv_q) < 0.01:
                    restated = True
                elif abs(qty - csv_q) > 0.01:
                    warn(f"{d}: portföljens antal {qty} stämmer inte med inköpscykeln {csv_q:.5f}")
            else:
                restated = vest_of.get(d, d) >= split_date
        if restated:
            qty, price = qty / SPLIT_RATIO, price * SPLIT_RATIO
        raw_purchase_price[d] = price
        restated_on[d] = restated
        if d in vest_of:  # purchased shares are locked until the matching shares vest
            restricted_lots.append((d, vest_of[d], qty))
        note = "EquatePlus " + r["Instrument"] + (" (omräknad till före split)" if restated else "")
        events.append(Event(d, "Köp", qty, qty * price, note=note))

    # Benefit-taxed shares (matching shares at delivery, free shares) and plan account.
    for r in portfolio:
        if r["Typ av instrument"] != "aktier" or r["Typ av insättning"] == "Köp":
            continue
        d = excel_date(r["Allokeringsdatum"])
        qty, price = float(r["Allokerad kvantitet"]), num(r["Lösenpris/kostnadsbas"])
        instr = r["Instrument"]
        if r["Plan"] == "Share Plan Account":
            msg = f"{d}: 'Share Plan Account' Tilldelning {qty:g} st @ {price} kr"
            if exclude_spa:
                warn(msg + " EJ medtagen (--exclude-share-plan-account).")
                continue
            warn(msg + " medtagen som Övr (behövs för att innehavet ska stämma med "
                 "utdelningarna) - kontrollera omkostnadsbeloppet.")
            note = "EquatePlus " + instr
            # The allocation is sometimes dated after the transfer of the same shares out
            # of EquatePlus; the shares must have been there by the transfer date.
            earlier = [excel_date(t["Datum"]) for t in transactions
                       if t.get("Ordertyp") == "Överför" and t.get("Status") == "Utförda"
                       and abs(float(t["Antal"]) - qty) < TOL
                       and 0 < (d - excel_date(t["Datum"])).days <= 3]
            if earlier:
                warn(f"{d}: 'Share Plan Account' daterad till {min(earlier)} då samma antal "
                     "överfördes ut från EquatePlus det datumet.")
                d = min(earlier)
                note += f" (daterad till överföringen, tilldelning {excel_date(r['Allokeringsdatum'])})"
            events.append(Event(d, "Övr", qty, qty * (price or 0), note=note))
            continue
        if price is None:
            warn(f"{d}: {instr} utan pris, hoppas över ({qty} st)")
            continue
        note = "EquatePlus " + instr
        if d < split_date and d in raw_purchase_price:
            ratio = raw_purchase_price[d] / price
            if 0.8 * SPLIT_RATIO < ratio < 1.2 * SPLIT_RATIO:
                warn(f"{d}: {instr} pris {price} är angivet efter split (köp samma dag "
                     f"{raw_purchase_price[d]:.2f}); multiplicerat med {SPLIT_RATIO}.")
                price *= SPLIT_RATIO
                note += ", pris justerat för split"
        if d in deliveries and abs(deliveries[d] - qty) > TOL:
            warn(f"{d}: {instr} {qty} st i portföljen men Leverans {deliveries[d]} st; "
                 "leveransens antal används.")
            qty = deliveries[d]
            note += ", antal enligt Leverans"
        events.append(Event(d, "Förm", qty, qty * price, note=note))

    # Actual delivery date of each matching grant: the Förm event with the same quantity
    # closest to the portfolio's vesting date ("Tillgänglig från" is sometimes off).
    delivered: dict[dt.date, dt.date] = {}
    forms = [e for e in events if e.kind == "Förm"]
    used: set[int] = set()
    for r in portfolio:
        if r["Typ av instrument"] != "begränsade aktieenheter" or not r["Tillgänglig från"]:
            continue
        alloc, vest = excel_date(r["Allokeringsdatum"]), excel_date(r["Tillgänglig från"])
        q = float(r["Allokerad kvantitet"])
        cands = [i for i, e in enumerate(forms) if i not in used and abs(e.qty - q) < 1e-4
                 and abs((e.date - vest).days) <= DELIVERY_WINDOW_DAYS]
        if cands:
            best = min(cands, key=lambda i: abs((forms[i].date - vest).days))
            used.add(best)
            delivered[alloc] = forms[best].date
    # Purchased shares are free when the matching shares are delivered, or at the latest
    # after the three-year holding period.
    restricted_lots = [(a, min(v, delivered.get(a, v), add_years(a, HOLDING_YEARS)), q)
                       for a, v, q in restricted_lots]

    by_date: dict[dt.date, set[float]] = {}
    for e in events:
        if e.kind == "Förm" and e.qty:
            by_date.setdefault(e.date, set()).add(round(e.cost / e.qty, 2))
    for d, prices in sorted(by_date.items()):
        if len(prices) > 1:
            warn(f"{d}: flera Förm-rader med olika pris {sorted(prices)} samma dag - "
                 "stäm av förmånsvärdet mot kontrolluppgiften.")

    transfers = [(excel_date(t["Datum"]), float(t["Antal"]), t["Orderreferens"])
                 for t in transactions
                 if t.get("Ordertyp") == "Överför" and t.get("Status") == "Utförda"
                 and t.get("Produkttyp", "").startswith("aktie")]
    dividends = [(excel_date(t["Datum"]), float(t["Antal"]))
                 for t in transactions
                 if t.get("Ordertyp") == "Utdelning" and t.get("Produkttyp") == "aktier"
                 and t.get("Status") == "Utförda"]
    if any(t.get("Ordertyp", "").startswith("Sälj") and t.get("Status") == "Utförda"
           for t in transactions):
        warn("Genomförda försäljningar i EquatePlus finns men hanteras inte av verktyget.")

    pending = [(excel_date(r["Allokeringsdatum"]), excel_date(r["Tillgänglig från"]),
                float(r["Utestående antal"] or 0))
               for r in portfolio
               if r["Typ av instrument"] == "begränsade aktieenheter"
               and float(r["Utestående antal"] or 0) > 0]
    current = sum(float(r["Utestående antal"] or 0) for r in portfolio
                  if r["Typ av instrument"] == "aktier")

    # Unvested matching units, in units valid on the allocation date (EquatePlus restated
    # the same allocations as the purchases).
    match_lots = []
    for r in portfolio:
        if r["Typ av instrument"] == "begränsade aktieenheter" and r["Tillgänglig från"]:
            alloc = excel_date(r["Allokeringsdatum"])
            q = float(r["Allokerad kvantitet"])
            if restated_on.get(alloc):
                q /= SPLIT_RATIO
            match_lots.append((alloc, delivered.get(alloc, excel_date(r["Tillgänglig från"])), q))

    as_of = None
    for r in raw_pf[:5]:
        if r.get("A", "").strip() == "Per den" and r.get("B"):
            try:
                as_of = sv_date(" ".join(r["B"].split()[:3]))
            except (KeyError, ValueError):
                pass
    restricted_now = sum(float(r["Utestående antal"] or 0) for r in portfolio
                         if r["Typ av instrument"] == "aktier" and r["Typ av insättning"] == "Köp")
    match_now = sum(p[2] for p in pending)
    return EquatePlus(events, transfers, dividends, pending, current, restricted_lots,
                      match_lots, as_of, restricted_now, match_now)


def equateplus_balance(ep: EquatePlus, split_date: dt.date, before: dt.date) -> float:
    """Shares held in EquatePlus just before `before` (units valid on that date)."""
    items = [(e.date, e.order, e.qty) for e in ep.events]
    items += [(d, 5, -q) for d, q, _ in ep.transfers]
    items.append((split_date, ORDER["S"], None))
    bal = 0.0
    for d, _, q in sorted(items, key=lambda x: (x[0], x[1])):
        if d >= before:
            break
        bal = bal * SPLIT_RATIO if q is None else bal + q
    return bal


# --------------------------------------------------------------------------- Nordnet


def load_nordnet(paths: list[str]) -> Nordnet:
    nn = Nordnet(files=paths)
    seen: set[str] = set()
    rows = []
    for p in paths:
        for r in read_nordnet(p):
            if not r.get("Värdepapper", "").upper().startswith(SECURITY_PREFIX):
                continue
            if r.get("Makuleringsdatum"):
                continue
            key = r.get("Id") or "|".join(r.values())
            if key in seen:
                continue
            seen.add(key)
            rows.append(r)
    if not rows:
        return nn

    def tdate(r):
        return dt.date.fromisoformat(r.get("Affärsdag") or r["Bokföringsdag"])

    def rid(r):
        try:
            return int(r.get("Id") or 0)
        except ValueError:
            return 0

    rows.sort(key=lambda r: (tdate(r), rid(r)))

    def delta(r) -> float | None:
        t = r["Transaktionstyp"].upper()
        q = num(r.get("Antal")) or 0.0
        if t in ("KÖPT", "INLÄGG VP"):
            return q
        if t in ("SÅLT", "UTTAG VP"):
            return -q
        return 0.0 if "SPLIT" not in t else None

    first = rows[0]
    nn.start_date = tdate(first)
    if "SPLIT INLÄGG" in first["Transaktionstyp"].upper():
        nn.opening = (num(first.get("Antal")) or 0) / SPLIT_RATIO
    else:
        after = num(first.get("Totalt antal"))
        d0 = delta(first)
        nn.opening = (after - d0) if after is not None and d0 is not None else 0.0

    holdings = nn.opening
    eod: dict[dt.date, float] = {}  # holdings at end of each trade date
    for r in rows:
        t = r["Transaktionstyp"].upper()
        d = tdate(r)
        q = num(r.get("Antal")) or 0.0
        text = r.get("Transaktionstext", "")
        if t == "SPLIT INLÄGG VP":
            nn.split_date = d
            holdings = q
        elif t == "SPLIT UTTAG VP":
            pass
        elif t == "SÅLT":
            amount = num(r.get("Belopp")) or 0.0
            fee = num(r.get("Total Avgift")) or 0.0
            price = num(r.get("Kurs")) or 0.0
            nn.events.append(Event(d, "Förs", q, proceeds=amount,
                                   note=f"Nordnet sålt {q:g} st à {price:g} kr, avgift {fee:g} kr",
                                   ref=r.get("Notanummer", "")))
            holdings -= q
        elif t == "KÖPT":
            amount = -(num(r.get("Belopp")) or 0.0)
            nn.events.append(Event(d, "Köp", q, amount, note=f"Nordnet köpt {q:g} st",
                                   ref=r.get("Notanummer", "")))
            holdings += q
        elif t == "OMVANDLING INLÄGG VP":
            cost = num(r.get("Inköpsvärde"))
            if cost is None:
                warn(f"Nordnet {d}: OMVANDLING {q:g} st utan inköpsvärde - omkostnadsbelopp 0 "
                     "används, ange rätt belopp manuellt.")
                cost = 0.0
            else:
                warn(f"Nordnet {d}: OMVANDLING {q:g} st ('{text}') medtagen som Övr med "
                     f"inköpsvärde {cost:g} kr enligt Nordnet - kontrollera ursprunget.")
            nn.events.append(Event(d, "Övr", q, cost,
                                   note=f"Nordnet omvandling {q:g} st, inköpsvärde {cost:g} kr"))
            holdings += q
        elif t == "OMVANDLING UTTAG VP":
            warn(f"Nordnet {d}: OMVANDLING UTTAG {q:g} st ('{text}') - hanteras inte, "
                 "kontrollera manuellt.")
            holdings -= q
        elif t == "INLÄGG VP":
            squashed = text.lower().replace(" ", "")
            if "transfer" in squashed or "computershare" in squashed or squashed.endswith("cs"):
                nn.transfers_in.append((d, q, text))
            else:
                warn(f"Nordnet {d}: INLÄGG VP {q:g} st ('{text}') från okänd källa - "
                     "ej med i beräkningen, lägg till manuellt om det är ett förvärv.")
            holdings += q
        elif t == "UTTAG VP":
            warn(f"Nordnet {d}: UTTAG VP {q:g} st ('{text}') - ej försäljning, "
                 "kontrollera vart aktierna flyttades.")
            holdings -= q
        elif t in ("UTDELNING", "PRELSKATT UTD"):
            pass
        else:
            warn(f"Nordnet {d}: okänd transaktionstyp '{r['Transaktionstyp']}' ignorerad.")
        tot = num(r.get("Totalt antal"))
        if tot is not None and t not in ("UTDELNING", "PRELSKATT UTD", "SPLIT UTTAG VP") \
                and abs(tot - holdings) > TOL:
            warn(f"Nordnet {d}: beräknat innehav {holdings:g} men 'Totalt antal' {tot:g}.")
            holdings = tot
        eod[d] = holdings
    nn.eod = eod
    nn.closing = holdings
    nn.closing_date = tdate(rows[-1])
    return nn


TRANSFER_WINDOW_DAYS = 60


def match_transfers(ep: EquatePlus, nn: Nordnet) -> float:
    """Pair EquatePlus transfers out with Nordnet transfers in (same qty, within
    TRANSFER_WINDOW_DAYS, earliest first). Returns the number of shares that left
    EquatePlus before the Nordnet export starts but arrived in Nordnet after (in transit)."""
    if not nn.start_date:
        return 0.0
    window = dt.timedelta(days=TRANSFER_WINDOW_DAYS)
    unmatched_in = sorted(range(len(nn.transfers_in)), key=lambda i: nn.transfers_in[i][0])
    in_transit = 0.0
    for d, q, ref in sorted(ep.transfers):
        if d < nn.start_date - window:
            continue
        hit = next((i for i in unmatched_in
                    if abs(nn.transfers_in[i][1] - q) < TOL
                    and dt.timedelta(0) <= nn.transfers_in[i][0] - d <= window), None)
        if hit is not None:
            unmatched_in.remove(hit)
            nn.matched_in.add(hit)
            nn.matched_out.add(ref)
            if d < nn.start_date:
                in_transit += q
        elif d >= nn.start_date:
            warn(f"EquatePlus-överföring {d} {q:g} st ({ref}) hittades inte i Nordnet.")
    for i in unmatched_in:
        d, q, text = nn.transfers_in[i]
        warn(f"Nordnet-inlägg {d} {q:g} st ('{text}') saknar motsvarande EquatePlus-överföring.")
    return in_transit


# --------------------------------------------------------------------------- calculation


def compute(events: list[Event], name: str):
    rows, sales = [], []
    total_qty = 0.0
    total_cost = 0.0
    for e in sorted(events, key=lambda e: (e.date, e.order)):
        if e.kind == "S":
            delta, change = total_qty * (SPLIT_RATIO - 1), 0.0
        elif e.kind in ("Förs", "Adj"):
            if e.qty > total_qty + TOL:
                warn(f"{e.date}: försäljning av {e.qty:g} st men endast {total_qty:g} st i "
                     "beräkningen - historiken är ofullständig.")
            avg_before = total_cost / total_qty if total_qty else 0.0
            delta, change = -e.qty, -e.qty * avg_before
            if e.kind == "Förs":
                sales.append({"Datum": e.date.isoformat(), "Antal": e.qty, "Beteckning": name,
                              "Försäljningspris": e.proceeds, "Omkostnadsbelopp": -change,
                              "Resultat": e.proceeds + change, "Nota": e.ref})
        else:
            delta, change = e.qty, e.cost
        total_qty += delta
        total_cost += change
        if e.kind == "Förs":
            sales[-1]["LedgerEfter"] = total_qty
        if abs(total_qty) < TOL:
            total_qty, total_cost = 0.0, 0.0
        avg = total_cost / total_qty if total_qty else 0.0
        rows.append({
            "Beteckning": name,
            "Datum": e.date.isoformat(),
            "Händelse": "Övr" if e.kind == "Adj" else e.kind,
            "Antal +/-": delta,
            "Totalt antal": total_qty,
            "Inköpspris (+) / utnyttjat omk.belopp (-), kr": change,
            "Totalt outnyttjat omk.belopp, kr": total_cost,
            "Genomsnittligt omk.belopp, kr": avg,
            "Kommentar": e.note,
        })
    return rows, sales


def ledger_total_before(events: list[Event], before: dt.date, kind_order: int) -> float:
    """Total shares in the ledger before (before, kind_order) - acquisitions/split/sales."""
    tot = 0.0
    for e in sorted(events, key=lambda e: (e.date, e.order)):
        if (e.date, e.order) >= (before, kind_order):
            break
        if e.kind == "S":
            tot *= SPLIT_RATIO
        elif e.kind in ("Förs", "Adj"):
            tot -= e.qty
        else:
            tot += e.qty
    return tot


# --------------------------------------------------------------------------- holdings log

# Same-day order: everything else before sales (as in the SKV 2198 table), transfers
# out of EquatePlus before arrivals in Nordnet.
LOG_ORDER = {"Adj": -1, "S": 0, "Grant": 1, "Vest": 1, "Köp": 1, "Förm": 2, "Övr": 2,
             "Out": 4, "In": 5, "Förs": 9}
LOG_LABEL = {"Adj": "Justering", "S": "Split", "Grant": "Tilldelning matchning",
             "Vest": "Intjänande", "Köp": "Köp", "Förm": "Förm", "Övr": "Övr",
             "Out": "Överföring ut", "In": "Överföring in", "Förs": "Försäljning"}


def build_log(ep: EquatePlus, nn: Nordnet, ledger_events: list[Event], sales: list[dict],
              split_date: dt.date, name: str) -> list[dict]:
    """One row per event that changes holdings anywhere, with all holdings after it."""
    cutoff = ep.as_of or max([e.date for e in ledger_events] + [nn.closing_date or dt.date.min])
    items: list[dict] = []

    def add(d, kind, qty, source, text, **extra):
        items.append({"date": d, "order": LOG_ORDER[kind], "seq": len(items), "kind": kind,
                      "qty": qty, "source": source, "text": text, **extra})

    locked = {a: q for a, _, q in ep.restricted_lots}
    for e in ep.events:
        add(e.date, e.kind, e.qty, "EquatePlus", e.note,
            locked=locked.get(e.date, 0.0) if e.kind == "Köp" else 0.0)
    vests: dict[dt.date, dict[str, list]] = {}
    for a, v, q in ep.match_lots:
        if a <= cutoff:
            add(a, "Grant", q, "EquatePlus", f"Företagsmatchning tilldelad, intjänas {v}")
        if v <= cutoff:
            vests.setdefault(v, {"r": [], "m": []})["m"].append((a, q))
    for a, v, q in ep.restricted_lots:
        if v <= cutoff:
            vests.setdefault(v, {"r": [], "m": []})["r"].append((a, q))
    for v, lots in vests.items():
        add(v, "Vest", 0.0, "EquatePlus", "", rlots=lots["r"], mlots=lots["m"])
    for d, q, ref in ep.transfers:
        add(d, "Out", q, "EquatePlus", f"Till Nordnet ({ref})", matched=ref in nn.matched_out)
    for i, (d, q, text) in enumerate(nn.transfers_in):
        add(d, "In", q, "Nordnet", f"Från EquatePlus ({text})", matched=i in nn.matched_in)
    for e in nn.events:
        add(e.date, e.kind, e.qty, "Nordnet", e.note)
    for e in ledger_events:
        if e.kind in ("S", "Adj"):
            add(e.date, e.kind, e.qty, "", e.note)
    items.sort(key=lambda x: (x["date"], x["order"], x["seq"]))

    def units(alloc, day, q):
        return q * (SPLIT_RATIO if alloc < split_date <= day else 1)

    ep_tot = restr = match = transit = nord = 0.0
    sale_iter = iter(sales)
    nord_eod: dict[dt.date, float] = {}
    rows = []
    for it in items:
        d, k, q, src = it["date"], it["kind"], it["qty"], it["source"]
        change, text, sale = q, it["text"], None
        before = ep_tot + transit + nord
        if k == "S":
            ep_tot, restr, match, transit, nord = (x * SPLIT_RATIO for x in
                                                   (ep_tot, restr, match, transit, nord))
            change = ep_tot + transit + nord - before
        elif k == "Grant":
            match += q
        elif k == "Vest":
            rq = sum(units(a, d, x) for a, x in it["rlots"])
            mq = sum(units(a, d, x) for a, x in it["mlots"])
            restr -= rq
            match -= mq
            change = rq
            text = (f"Begränsning upphör för {fmt_sv(rq, 5)} st köpta aktier; "
                    f"{fmt_sv(mq, 5)} st företagsmatchning intjänad (levereras som Förm)")
        elif src == "EquatePlus" and k in ("Köp", "Förm", "Övr"):
            ep_tot += q
            restr += it.get("locked", 0.0)
        elif k == "Out":
            ep_tot -= q
            transit += q
            change = -q
            if not it["matched"]:
                text += " - ej hittad i Nordnet"
        elif k == "In":
            nord += q
            if it["matched"]:
                transit -= q
            else:
                text += " - ingen motsvarande EquatePlus-överföring"
        elif src == "Nordnet" and k in ("Köp", "Övr"):
            nord += q
        elif k == "Förs":
            nord -= q
            change = -q
            sale = next(sale_iter)
        elif k == "Adj":
            transit -= q
            change = -q
        free = ep_tot - restr
        total = ep_tot + transit + nord
        if free < -0.01:
            warn(f"{d}: fler begränsade aktier ({restr:.5f}) än innehav i EquatePlus ({ep_tot:.5f}).")
        if sale and abs(total - sale["LedgerEfter"]) > 0.01:
            warn(f"{d}: innehav efter försäljning {total:.5f} st stämmer inte med "
                 f"SKV 2198-tabellens {sale['LedgerEfter']:.5f} st.")
        nord_eod[d] = nord
        res = sale["Resultat"] if sale else None
        rows.append({
            "Inkomstår": str(d.year), "Datum": d.isoformat(), "Händelse": LOG_LABEL[k],
            "Källa": src, "Antal": change, "Beteckning": name,
            "Försäljningspris": sale["Försäljningspris"] if sale else "",
            "Omkostnadsbelopp": sale["Omkostnadsbelopp"] if sale else "",
            "Vinst": max(res, 0.0) if sale else "", "Förlust": max(-res, 0.0) if sale else "",
            "Nota": sale["Nota"] if sale else "", "Kommentar": text,
            HOLD_FREE: free, HOLD_RESTR: restr, HOLD_MATCH: match, "Under överföring": transit,
            "Kvar Nordnet": nord, "Kvar totalt": total,
        })

    # Cross-checks of the simulated holdings.
    for d, h in sorted(nn.eod.items()):
        if d in nord_eod and abs(nord_eod[d] - h) > TOL:
            warn(f"Innehavslogg {d}: Nordnet {nord_eod[d]:g} st, enligt Nordnet {h:g} st.")
    if ep.as_of:
        for label, calc, real in (("EquatePlus totalt", ep_tot, ep.current),
                                  ("begränsade", restr, ep.restricted_now),
                                  ("ej intjänad företagsmatchning", match, ep.match_now)):
            if abs(calc - real) > 0.01:
                warn(f"Innehavslogg {ep.as_of}: {label} {calc:.5f} st, enligt portföljen {real:.5f} st.")
    if nn.start_date and abs(nord - nn.closing) > TOL:
        warn(f"Innehavslogg: Nordnet slutsaldo {nord:g} st, enligt Nordnet {nn.closing:g} st.")
    return rows


# --------------------------------------------------------------------------- output


def fmt_sv(v, decimals: int) -> str:
    return f"{v:,.{decimals}f}".replace(",", " ").replace(".", ",")


def write_csv(path, rows, decimals: dict[str, int], swedish: bool) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";" if swedish else ",")
        w.writeheader()
        for r in rows:
            out = {}
            for k, v in r.items():
                if k in decimals and isinstance(v, (int, float)):
                    out[k] = fmt_sv(v, decimals[k]).replace(" ", "") if swedish \
                        else f"{v:.{decimals[k]}f}"
                else:
                    out[k] = v
            w.writerow(out)


def print_table(headers, data, left=()) -> None:
    widths = [max(len(h), *(len(d[i]) for d in data)) if data else len(h)
              for i, h in enumerate(headers)]
    line = lambda cells: "  ".join(c.ljust(w) if i in left else c.rjust(w)
                                   for i, (c, w) in enumerate(zip(cells, widths)))
    print(line(headers))
    print("  ".join("-" * w for w in widths))
    for d in data:
        print(line(d))


SKV_DEC = {"Antal +/-": 5, "Totalt antal": 5, "Inköpspris (+) / utnyttjat omk.belopp (-), kr": 2,
           "Totalt outnyttjat omk.belopp, kr": 2, "Genomsnittligt omk.belopp, kr": 2}
K4_DEC = {"Antal": 5, "Försäljningspris": 0, "Omkostnadsbelopp": 0, "Vinst": 0, "Förlust": 0}
HOLD_FREE = "EQP Tillgänglig för handel"
HOLD_RESTR = "EQP Tillgänglig med begränsningar"
HOLD_MATCH = "EQP Företagsmatchning (ej intjänad, ingår ej i totalt)"
LOG_DEC = {"Antal": 5, "Försäljningspris": 0, "Omkostnadsbelopp": 0, "Vinst": 0, "Förlust": 0,
           HOLD_FREE: 5, HOLD_RESTR: 5, HOLD_MATCH: 5, "Under överföring": 5,
           "Kvar Nordnet": 5, "Kvar totalt": 5}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0],
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("directory", nargs="?", default=".", help="folder with the exports")
    ap.add_argument("-o", "--output", default="skv2198_saab.csv", help="SKV 2198 CSV (%(default)s)")
    ap.add_argument("--k4", default="k4_saab.csv", help="K4 sales CSV (%(default)s)")
    ap.add_argument("--k4-holdings", default="k4_innehav_saab.csv",
                    help="holdings log CSV, every event with holdings after it (%(default)s)")
    ap.add_argument("--year", type=int, help="show K4 sales only for this income year")
    ap.add_argument("--name", default="Saab B", help="Beteckning (%(default)s)")
    ap.add_argument("--nordnet", nargs="*", metavar="CSV",
                    help="Nordnet export file(s); default: Nordnet*.csv in the folder")
    ap.add_argument("--split-date", type=dt.date.fromisoformat,
                    help="date of the 4:1 split, YYYY-MM-DD (default: from Nordnet, "
                         f"else {FALLBACK_SPLIT_DATE})")
    ap.add_argument("--exclude-share-plan-account", action="store_true",
                    help="leave out the 'Share Plan Account' allocation")
    ap.add_argument("--no-adjust", action="store_true",
                    help="do not insert an adjustment row for shares unaccounted for")
    ap.add_argument("--decimal-point", action="store_true",
                    help="write CSV with '.' decimals and ',' separator")
    args = ap.parse_args()

    nn_paths = args.nordnet if args.nordnet is not None else find(args.directory, "Nordnet*.csv")
    nn = load_nordnet(nn_paths) if nn_paths else Nordnet()
    split_date = args.split_date or nn.split_date or FALLBACK_SPLIT_DATE
    if not args.split_date and not nn.split_date:
        warn(f"Splitdatum okänt, antar {split_date} (ange --split-date).")

    ep = load_equateplus(args.directory, split_date, args.exclude_share_plan_account)
    events = list(ep.events) + list(nn.events)
    if ep.events and min(e.date for e in ep.events) < split_date:
        events.append(Event(split_date, "S", note=f"Split {SPLIT_RATIO}:1"))

    # Check the restricted / unvested split against the portfolio per its "Per den" date.
    if ep.as_of:
        r_calc = lots_held(ep.restricted_lots, split_date, ep.as_of)
        m_calc = lots_held(ep.match_lots, split_date, ep.as_of)
        if abs(r_calc - ep.restricted_now) > 0.01:
            warn(f"Begränsade aktier {ep.as_of}: beräknat {r_calc:.5f} st, enligt portföljen "
                 f"{ep.restricted_now:.5f} st.")
        if abs(m_calc - ep.match_now) > 0.01:
            warn(f"Ej intjänad företagsmatchning {ep.as_of}: beräknat {m_calc:.5f} st, enligt "
                 f"portföljen {ep.match_now:.5f} st.")

    # Check EquatePlus holdings against the dividend share counts.
    for d, q in ep.dividend_holdings:
        b = equateplus_balance(ep, split_date, d)
        if abs(b - q) > 0.01:
            warn(f"EquatePlus innehav {d}: beräknat {b:.5f} st, enligt utdelning {q:.5f} st.")

    # Reconcile with Nordnet opening balance; adjust for shares that left before the export.
    adj_qty = 0.0
    if nn.start_date:
        in_transit = match_transfers(ep, nn)
        start = nn.start_date
        ledger = ledger_total_before(events, start, ORDER["Adj"])
        outside = ledger - equateplus_balance(ep, split_date, start) - in_transit
        adj_qty = outside - nn.opening
        if adj_qty > 0.01 and not args.no_adjust:
            events.append(Event(start, "Adj", adj_qty,
                                note="Justering: aktier avyttrade/flyttade före Nordnet-exportens "
                                     "början (okänd tidpunkt)"))
            warn(f"{adj_qty:.5f} st (räknat i antal före split om datumet ligger före split) "
                 f"som förts ut från EquatePlus finns inte i Nordnet vid exportens början "
                 f"{start}. De tas bort som 'Övr' till genomsnittligt omkostnadsbelopp på "
                 "det datumet. Ta med en Nordnet-export som täcker hela historiken (även "
                 "gamla ISIN SE0000112385) eller redovisa tidigare försäljningar för ett exakt "
                 "resultat.")
        elif adj_qty < -0.01:
            warn(f"Nordnet hade {-adj_qty:.5f} st fler aktier vid exportens början än "
                 "EquatePlus förklarar - förvärv från annan källa saknas i beräkningen.")

    rows, sales = compute(events, args.name)

    print(f"SKV 2198 - Genomsnittligt omkostnadsbelopp - {args.name}\n")
    print_table(
        ["Datum", "Händ.", "Antal +/-", "Totalt antal", "Inköpspris/utnyttjat",
         "Totalt omk.belopp", "Genomsn. omk.belopp", "Kommentar"],
        [[r["Datum"], r["Händelse"], fmt_sv(r["Antal +/-"], 5), fmt_sv(r["Totalt antal"], 5),
          fmt_sv(r["Inköpspris (+) / utnyttjat omk.belopp (-), kr"], 2),
          fmt_sv(r["Totalt outnyttjat omk.belopp, kr"], 2),
          fmt_sv(r["Genomsnittligt omk.belopp, kr"], 2), r["Kommentar"]] for r in rows],
        left={0, 1, 7})
    write_csv(args.output, rows, SKV_DEC, not args.decimal_point)
    print(f"\nSKV 2198-tabell sparad: {args.output}")

    # K4
    if sales:
        shown = [s for s in sales if args.year is None or s["Datum"][:4] == str(args.year)]
        k4_rows = []
        years = sorted({s["Datum"][:4] for s in shown})
        print("\nK4 avsnitt A - försäljningar")
        data = []
        for y in years:
            ys = [s for s in shown if s["Datum"][:4] == y]
            for s in ys:
                res = s["Resultat"]
                k4_rows.append({"Inkomstår": y, "Datum": s["Datum"], "Antal": s["Antal"],
                                "Beteckning": s["Beteckning"],
                                "Försäljningspris": s["Försäljningspris"],
                                "Omkostnadsbelopp": s["Omkostnadsbelopp"],
                                "Vinst": max(res, 0.0), "Förlust": max(-res, 0.0),
                                "Nota": s["Nota"]})
                data.append([y, s["Datum"], fmt_sv(s["Antal"], 0), fmt_sv(s["Försäljningspris"], 0),
                             fmt_sv(s["Omkostnadsbelopp"], 0), fmt_sv(max(res, 0), 0),
                             fmt_sv(max(-res, 0), 0)])
            tq = sum(s["Antal"] for s in ys)
            tp = sum(s["Försäljningspris"] for s in ys)
            to = sum(s["Omkostnadsbelopp"] for s in ys)
            tr = tp - to
            k4_rows.append({"Inkomstår": y, "Datum": "SUMMA", "Antal": tq,
                            "Beteckning": args.name, "Försäljningspris": tp,
                            "Omkostnadsbelopp": to, "Vinst": max(tr, 0.0),
                            "Förlust": max(-tr, 0.0), "Nota": ""})
            data.append([y, "SUMMA K4", fmt_sv(tq, 0), fmt_sv(tp, 0), fmt_sv(to, 0),
                         fmt_sv(max(tr, 0), 0), fmt_sv(max(-tr, 0), 0)])
            data.append([""] * 7)
        print_table(["År", "Datum", "Antal", "Försäljningspris", "Omkostnadsbelopp",
                     "Vinst", "Förlust"], data[:-1], left={0, 1})
        if k4_rows:
            write_csv(args.k4, k4_rows, K4_DEC, not args.decimal_point)
            print(f"\nK4-underlag sparat: {args.k4}  (avrunda till hela kronor på blanketten)")

    # Holdings log
    log = build_log(ep, nn, events, sales, split_date, args.name)
    if args.year is not None:
        log = [r for r in log if r["Inkomstår"] == str(args.year)]
    if log:
        write_csv(args.k4_holdings, log, LOG_DEC, not args.decimal_point)
        print(f"Innehavslogg (alla händelser, innehav efter varje) sparad: {args.k4_holdings}")

    # Summary
    last = rows[-1]
    print("\nSammanfattning")
    print(f"  Totalt innehav enligt beräkningen:  {fmt_sv(last['Totalt antal'], 5)} st, "
          f"genomsnitt {fmt_sv(last['Genomsnittligt omk.belopp, kr'], 2)} kr/st")
    held = ep.current + (nn.closing if nn.start_date else 0.0)
    print(f"  Faktiskt innehav: EquatePlus {fmt_sv(ep.current, 5)} st"
          + (f" + Nordnet {fmt_sv(nn.closing, 0)} st ({nn.closing_date})" if nn.start_date else "")
          + f" = {fmt_sv(held, 5)} st")
    if abs(held - last["Totalt antal"]) > 0.01:
        warn(f"Beräknat innehav {last['Totalt antal']:.5f} st skiljer sig från faktiskt "
             f"{held:.5f} st.")
    if ep.pending:
        print(f"  Ej intjänade matchningsaktier (ingår ej): "
              f"{fmt_sv(sum(p[2] for p in ep.pending), 5)} st, intjänas "
              f"{min(p[1] for p in ep.pending)} - {max(p[1] for p in ep.pending)}")
    if not nn.start_date:
        print("  Ingen Nordnet-fil: inga försäljningar med i beräkningen.")

    if warnings:
        print("\nVarningar / kontrollera:")
        for m in warnings:
            print("  * " + m)
    print("\nFörmånsvärdet för matchningsaktier bör stämmas av mot kontrolluppgift/lönebesked.")


if __name__ == "__main__":
    main()
