#!/usr/bin/env bash
# Build eqpskv for Linux, macOS and Windows (amd64 + arm64).
# Output: dist/eqpskv-<os>-<arch>[.exe]
# Usage: ./build.sh
set -euo pipefail

cd "$(dirname "$0")"

# Find the Go toolchain (also when /usr/local/go/bin is not on PATH).
if ! command -v go >/dev/null 2>&1; then
    if [ -x /usr/local/go/bin/go ]; then
        export PATH="$PATH:/usr/local/go/bin"
    else
        echo "error: go not found - install Go or add it to PATH" >&2
        exit 1
    fi
fi

OUT=dist
TARGETS=(
    linux/amd64
    linux/arm64
    darwin/amd64
    darwin/arm64
    windows/amd64
    windows/arm64
)

go vet ./...

rm -rf "$OUT"
mkdir -p "$OUT"

for target in "${TARGETS[@]}"; do
    os="${target%/*}"
    arch="${target#*/}"
    ext=""
    [ "$os" = "windows" ] && ext=".exe"
    bin="$OUT/eqpskv-$os-$arch$ext"
    echo "building $bin"
    CGO_ENABLED=0 GOOS="$os" GOARCH="$arch" \
        go build -trimpath -ldflags "-s -w" -o "$bin" .
done

# Checksums
if command -v sha256sum >/dev/null 2>&1; then
    (cd "$OUT" && sha256sum eqpskv-* > SHA256SUMS)
elif command -v shasum >/dev/null 2>&1; then
    (cd "$OUT" && shasum -a 256 eqpskv-* > SHA256SUMS)
fi

echo "done:"
ls -l "$OUT"
