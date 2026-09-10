#!/usr/bin/env sh
# ac2f pack - kaynak dosyalari CorelDRAW VBA icin hazirlar.
#
# 1.2.0'dan itibaren kaynak saf ASCII'dir, bu yuzden kod sayfasi donusumu
# ZORUNLU DEGILDIR; src/*.bas dogrudan ice aktarilabilir. Bu betik yalnizca
# CRLF satir sonlu kopya uretmek icin kolaylik saglar. AC2F_CODEPAGE ile
# baska bir kod sayfasi istenebilir (varsayilan: ASCII uyumlu WINDOWS-1254).
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
