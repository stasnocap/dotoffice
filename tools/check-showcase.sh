#!/usr/bin/env bash
# Checks the phase 3 features of the converted showcase (see .github/workflows/ci.yml).
# Usage: tools/check-showcase.sh showcase.pdf convert.log
set -euo pipefail
pdf=$1 log=$2
fail() { echo "FAIL: $*"; exit 1; }
page_text() { pdftotext -f "$1" -l "$1" "$pdf" -; }

# Features that used to be skipped must not warn any more.
if grep -E "warning (docx\.columns|docx\.notes|docx\.floating)" "$log"; then fail "unsupported-feature warning"; fi

pdftotext "$pdf" all.txt
for text in "Contents" "8. Multiple columns" "Three balanced columns end" "Back to a single column" \
            "9. Footnotes and endnotes" "Footnote text uses the FootnoteText style." "Endnotes are numbered" \
            "10. Floating images" "Square wrapping" "11. Right-to-left text" "12. Fields and links" \
            "in roman numerals: xii" "«Customer»" "UAX #9"; do
  grep -qF "$text" all.txt || fail "missing: $text"
done

# Every TOC entry's page number must be the page its heading is on.
grep -E '^[0-9]+\. .* \.{3,} [0-9]+$' all.txt > toc.txt || fail "no TOC entries"
[ "$(wc -l < toc.txt)" -eq 12 ] || fail "expected 12 TOC entries"
while read -r line; do
  title=$(sed -E 's/ \.{3,} [0-9]+$//' <<< "$line"); page=$(grep -oE '[0-9]+$' <<< "$line")
  page_text "$page" | grep -qxF "$title" || fail "TOC says '$title' is on page $page"
done < toc.txt

# The PAGEREF in the fields paragraph points at the footnote section's page.
ref=$(grep -oE 'starts on page [0-9]+' all.txt | grep -oE '[0-9]+$')
page_text "$ref" | grep -qxF "9. Footnotes and endnotes" || fail "PAGEREF gave page $ref"

# Footnotes sit on the same page as their reference.
pages=$(pdfinfo "$pdf" | awk '/^Pages:/ {print $2}')
for p in $(seq 1 "$pages"); do
  if page_text "$p" | grep -qF "Footnotes are numbered automatically"; then
    page_text "$p" | grep -qF "Footnote text uses the FootnoteText style." || fail "footnote 1 not on page $p"
    page_text "$p" | grep -qF "A second footnote on the same page." || fail "footnote 2 not on page $p"
    found=1
  fi
done
[ "${found:-}" = 1 ] || fail "footnote reference not found"

# Hebrew and Arabic must round-trip through text extraction (needs the Noto fonts).
families=$(fc-list : family) # captured first: grep -q on a pipe would fail it under pipefail
if grep -q "Noto Sans Arabic" <<< "$families" && grep -q "Noto Sans Hebrew" <<< "$families"; then
  for text in "זהו טקסט בעברית" "מילים בעברית" "مرحبا بالعالم" "هذا نص عربي" "كلمات عربية"; do
    grep -qF "$text" all.txt || fail "right-to-left text did not round-trip: $text"
  done
  pdffonts "$pdf" > fonts.txt
  grep -q NotoSansArabic fonts.txt || fail "Noto Sans Arabic not embedded"
else
  echo "Noto Sans Arabic/Hebrew not installed: skipping right-to-left text checks"
fi
echo "showcase checks passed"
