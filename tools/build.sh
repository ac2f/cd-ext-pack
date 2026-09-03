#!/usr/bin/env sh
# ac2f pack - kaynak dosyalari CorelDRAW VBA icin hazirlar.
#
# src/*.bas depoda UTF-8 + LF tutulur. VBE "File > Import File" ANSI bekler;
# Turkce Windows'ta bu Windows-1254'tur. Bu betik build/ altina
# Windows-1254 + CRLF kopyalar.
set -eu

ROOT=$(CDPATH='' cd -- "$(dirname -- "$0")/.." && pwd)
SRC="$ROOT/src"
OUT="$ROOT/build"
CP="${AC2F_CODEPAGE:-WINDOWS-1254}"

mkdir -p "$OUT"

for f in "$SRC"/*.bas; do
    name=$(basename "$f")
    sed 's/$/\r/' "$f" | iconv -f UTF-8 -t "$CP" > "$OUT/$name"
    printf '  %-20s -> %s\n' "$name" "$OUT/$name"
done

echo
echo "Hazir. build/ altindaki .bas dosyalarini VBE > File > Import File ile aktarin."
