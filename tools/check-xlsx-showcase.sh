#!/usr/bin/env bash
# Checks the converted XLSX showcase (see .github/workflows/ci.yml). Page numbers and counts are derived from the PDF,
# never hardcoded: only stable markers (sheet names, formatted values, repeated title rows) are looked for.
# Usage: tools/check-xlsx-showcase.sh showcase-xlsx.pdf convert-xlsx.log
set -euo pipefail
pdf=$1 log=$2
fail() { echo "FAIL: $*"; exit 1; }
page_text() { pdftotext -layout -f "$1" -l "$1" "$pdf" -; }
header_of() { page_text "$1" | grep -m1 -E '\S' | sed -E 's/^\s+//; s/\s+$//'; }

pages=$(pdfinfo "$pdf" | awk '/^Pages:/ {print $2}')
[ "$pages" -ge 4 ] || fail "expected several pages, got $pages"
pdftotext -layout "$pdf" all.txt

# The footer counts pages across the whole workbook: every "Page N of M" has M = total, and page p says "Page p of M".
grep -qF "Page $pages of $pages" all.txt || fail "missing: Page $pages of $pages"
bad=$(grep -oE 'Page [0-9]+ of [0-9]+' all.txt | grep -vE " of $pages\$" || true)
[ -z "$bad" ] || fail "footer total is not $pages: $bad"
for p in $(seq 1 "$pages"); do
  [[ $(page_text "$p") =~ Page\ $p\ of\ $pages ]] || fail "page $p does not say 'Page $p of $pages'"
done

# Every visible sheet shows up as a page header; the hidden sheet is never printed.
for sheet in "Formats" "Long table" "Wide" "Text" "Rich text" "Rotation" "Gradients" "Pictures" "Conditional" "Paper"; do
  found=0
  for p in $(seq 1 "$pages"); do [ "$(header_of "$p")" = "$sheet" ] && found=1; done
  [ "$found" = 1 ] || fail "no page has the header '$sheet'"
done
if grep -q "HIDDENSHEETMARKER" all.txt; then fail "the hidden sheet was printed"; fi
if grep -E "warning xlsx\.formula-no-cache" "$log"; then fail "cached values were not used"; fi

# Number and date formats.
for text in "Number and date formats" "1,234.50" "25%" "25.67%" "2023-03-15" "Wednesday, March 15, 2023" "12:00 PM" "(1,234,567)" "1.23E+04" "#N/A" "2 1/2" "3 14/99" "2 3/8"; do
  grep -qF -- "$text" all.txt || fail "missing formatted value: $text"
done
grep -qE '^\s*Too wide for its column:\s+#+\s*$' all.txt || fail "an overly wide number should show hashes"
grep -qF "until a value stops it" all.txt || fail "overflowing text was cut off"

# Long table: the title and header rows repeat on every page of that sheet, and the manual break is honored.
long_pages=()
for p in $(seq 1 "$pages"); do [ "$(header_of "$p")" = "Long table" ] && long_pages+=("$p"); done
[ "${#long_pages[@]}" -ge 2 ] || fail "the long table should span several pages"
for p in "${long_pages[@]}"; do
  text=$(page_text "$p")
  for marker in "Long table report" "Item" "Name" "Amount"; do
    grep -qF "$marker" <<< "$text" || fail "title row text '$marker' missing on page $p"
  done
done
page_of() { for p in "${long_pages[@]}"; do [[ $(page_text "$p") == *"$1"* ]] && { echo "$p"; return; }; done; echo 0; }
last_page=$(page_of "Row 200 of the long table"); [ "$last_page" != 0 ] || fail "the last row is missing"
[ "$last_page" = "${long_pages[-1]}" ] || fail "the last row is not on the last page of the sheet"
before=$(page_of "Row 58 of the long table"); after=$(page_of "Row 59 of the long table")
[ "$before" != 0 ] && [ "$after" != 0 ] || fail "rows around the manual break are missing"
[ "$before" -lt "$after" ] || fail "the manual break after row 58 did not start a new page"
first_row=$(page_text "$after" | grep -oE 'Row [0-9]+ of the long table' | head -1)
[ "$first_row" = "Row 59 of the long table" ] || fail "the page after the break starts with '$first_row'"

# Wide: fit-to-width puts the first and the last column on the same landscape page.
first_col=$(for p in $(seq 1 "$pages"); do [[ $(page_text "$p") == *"Col 01"* ]] && echo "$p" || true; done)
last_col=$(for p in $(seq 1 "$pages"); do [[ $(page_text "$p") == *"Col 30"* ]] && echo "$p" || true; done)
[ -n "$first_col" ] && [ "$first_col" = "$last_col" ] || fail "'Col 01' (page ${first_col:-none}) and 'Col 30' (page ${last_col:-none}) are not on the same page"
read -r width height < <(pdfinfo -f "$first_col" -l "$first_col" "$pdf" | awk '/size:/ {print $4, $6; exit}')
awk -v w="$width" -v h="$height" 'BEGIN { exit !(w > h) }' || fail "the wide sheet is not landscape (${width}x${height})"

# Rich text: runs keep their text in order, in one line, and none of the features warns.
grep -qE 'Bold, italic, underlined, struck and red' all.txt || fail "rich text runs were not laid out in one line"
grep -qE 'H2O and E = mc2' all.txt || fail "sub/superscript runs are missing"
grep -qF "the runs flow" all.txt || fail "wrapped rich text is missing"
if grep -E "warning xlsx\.rich-text" "$log"; then fail "rich text should not warn any more"; fi

# Rotation: the sheet is laid out (pdftotext cannot read turned text reliably, so only the upright text is checked) and nothing warns.
grep -qF "Rotated and stacked text" all.txt || fail "the rotation sheet is missing"
grep -qF "Rows grow to fit rotated headings" all.txt || fail "the row under the rotated headings is missing"
if grep -E "warning xlsx\.rotation" "$log"; then fail "rotated text should not warn any more"; fi

# Gradients: the text on top of them is there, they are real shadings (not skipped) and nothing warns.
for text in "Left to right" "Top to bottom" "Path from the center"; do grep -qF -- "$text" all.txt || fail "missing gradient cell: $text"; done
if grep -E "warning xlsx\.gradient-fill" "$log"; then fail "gradient fills should not warn any more"; fi
grep -aqE '/ShadingType 2' "$pdf" || fail "no linear shading in the PDF"

# Pictures: placed by anchors, one of them continues across the manual page break, and only the shape is reported as unsupported.
pic_pages=()
for p in $(seq 1 "$pages"); do [ "$(header_of "$p")" = "Pictures" ] && pic_pages+=("$p"); done
[ "${#pic_pages[@]}" -ge 2 ] || fail "the pictures sheet should span two pages (manual break)"
for p in "${pic_pages[@]}"; do
  n=$(pdfimages -list -f "$p" -l "$p" "$pdf" | awk 'NR > 2' | wc -l)
  [ "$n" -ge 1 ] || fail "page $p of the pictures sheet has no picture"
done
first_n=$(pdfimages -list -f "${pic_pages[0]}" -l "${pic_pages[0]}" "$pdf" | awk 'NR > 2' | wc -l)
[ "$first_n" -ge 5 ] || fail "expected at least five pictures on the first page of the pictures sheet, found $first_n"
grep -E "warning xlsx\.drawing" "$log" | grep -qi "shape" || fail "the shape in the drawing should still be reported"
if grep -E "warning xlsx\.drawing" "$log" | grep -qi "picture\b.*not rendered\|images"; then fail "pictures must not be reported as unsupported"; fi

# Conditional formatting: the rules are applied (not reported), and what they draw is in the page content (fills, curves for the
# round icons). pdftotext cannot see colors, so the content streams are decompressed with qpdf and counted.
for text in "Conditional formatting" "Color scale" "Solid bar" "Symbols" "stopIfTrue"; do grep -qF -- "$text" all.txt || fail "missing conditional text: $text"; done
if grep -E "warning xlsx\.conditional-formatting" "$log"; then fail "supported conditional formatting should not warn"; fi
qpdf --qdf --object-streams=disable "$pdf" conditional-qdf.pdf
zebra=$(grep -ac '^0.8667 0.9216 0.9686 rg$' conditional-qdf.pdf || true)
[ "$zebra" -ge 5 ] || fail "expected five zebra stripes from the expression rule, found $zebra"
bad=$(grep -ac '^1 0.7804 0.8078 rg$' conditional-qdf.pdf || true)
[ "$bad" -ge 4 ] || fail "expected four cellIs fills, found $bad"
curves=$(grep -ac ' c$' conditional-qdf.pdf || true)
[ "$curves" -ge 80 ] || fail "expected at least 20 round icons (four curves each), found $curves curves"

# Paper sizes: the "Paper" sheet is A3 (842 x 1191 pt), every other sheet is A4 or Letter.
for p in $(seq 1 "$pages"); do
  if [ "$(header_of "$p")" = "Paper" ]; then
    pdfinfo -f "$p" -l "$p" "$pdf" | grep -E "^Page +$p size:" | grep -qE "(841|842)\.[0-9]+ x (1190|1191)" || fail "the Paper sheet should be A3"
  fi
done
