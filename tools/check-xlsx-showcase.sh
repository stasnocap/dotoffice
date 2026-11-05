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
for sheet in "Formats" "Long table" "Wide" "Text" "Rich text" "Rotation" "Gradients" "Pictures" "Conditional" "Paper" "Areas" "Comments" "Shapes" "Charts"; do
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
[ "$first_n" -ge 7 ] || fail "expected at least seven pictures on the first page of the pictures sheet, found $first_n"
qpdf --qdf --object-streams=disable "$pdf" pictures-qdf.pdf
[ "$(pdftotext -layout "$pdf" - | grep -c "purple, not red")" -ge 1 ] || fail "missing the indexed-color note"
grep -aq "^0.4392 0.1882 0.6275 rg$" pictures-qdf.pdf || fail "[Color3] should use the workbook palette (purple)"
grep -aqE "^-1 0 0 1 [0-9.]+ 0 cm$" pictures-qdf.pdf || fail "no horizontally mirrored picture (matrix -1 0 0 1)"
grep -aqE "^1 0 0 -1 0 [0-9.]+ cm$" pictures-qdf.pdf || fail "no vertically mirrored picture (matrix 1 0 0 -1)"
grep -E "warning xlsx\.drawing" "$log" | grep -qi "shape" || fail "the shape in the drawing should still be reported"
if grep -E "warning xlsx\.drawing" "$log" | grep -qi "picture\b.*not rendered\|images"; then fail "pictures must not be reported as unsupported"; fi

# Conditional formatting: the rules are applied (not reported), and what they draw is in the page content (fills, curves for the
# round icons). pdftotext cannot see colors, so the content streams are decompressed with qpdf and counted.
for text in "Conditional formatting" "Color scale" "Solid bar" "Symbols" "stopIfTrue" "More icon sets" "5Quarters" "4RedToBlack" "Excel 2010 extension rules" "5Boxes"; do grep -qF -- "$text" all.txt || fail "missing conditional text: $text"; done
if grep -E "warning xlsx\.conditional-formatting" "$log"; then fail "supported conditional formatting should not warn"; fi
qpdf --qdf --object-streams=disable "$pdf" conditional-qdf.pdf
zebra=$(grep -ac '^0.8667 0.9216 0.9686 rg$' conditional-qdf.pdf || true)
[ "$zebra" -ge 5 ] || fail "expected five zebra stripes from the expression rule, found $zebra"
bad=$(grep -ac '^1 0.7804 0.8078 rg$' conditional-qdf.pdf || true)
[ "$bad" -ge 5 ] || fail "expected four cellIs fills plus the tomorrow cell, found $bad"
yellow=$(grep -ac '^1 0.851 0.4 rg$' conditional-qdf.pdf || true)
[ "$yellow" -ge 1 ] || fail "the time-period rule for today should have filled a cell (found $yellow)"
grep -aq '^0.251 0.251 0.251 rg$' conditional-qdf.pdf || fail "the level icons (quarters, rating bars) are missing"
grep -aq '^0.9569 0.6275 0.6275 rg$' conditional-qdf.pdf || fail "the pink icon of 4RedToBlack is missing"
grep -aq '^0.8941 0.8745 0.9255 rg$' conditional-qdf.pdf || fail "the x14-only cellIs rule (inline dxf) did not fill a cell"
grep -aq '^0.1843 0.3333 0.5922 rg$' conditional-qdf.pdf || fail "the data bars should have a blue border"
grep -aq '^0.498 0 0 rg$' conditional-qdf.pdf || fail "the negative data bars should have a dark red border"
curves=$(grep -ac ' c$' conditional-qdf.pdf || true)
[ "$curves" -ge 80 ] || fail "expected at least 20 round icons (four curves each), found $curves curves"

# Shapes: drawn as vector content on a single page (themed blue fill, group parts, a dashed outline), with their text, and not reported.
shape_pages=0; shape_page=0
for p in $(seq 1 "$pages"); do [ "$(header_of "$p")" = "Shapes" ] && { shape_pages=$((shape_pages + 1)); shape_page=$p; }; done
[ "$shape_pages" = 1 ] || fail "the shapes sheet should fit one page, found $shape_pages"
shape_text=$(page_text "$shape_page")
for text in "Theme style" "Rounded rectangle" "Dashed oval" "A text box has no fill" "Hexagon"; do
  [[ $shape_text == *"$text"* ]] || fail "missing shape text: $text"
done
qpdf --qdf --object-streams=disable "$pdf" shapes-qdf.pdf
for color in "0.2667 0.4471 0.7686" "0.9294 0.4902 0.1922" "0.4392 0.1882 0.6275" "0.7529 0 0"; do
  grep -aq "^$color rg$" shapes-qdf.pdf || fail "shape fill $color is missing"
done
grep -aq '^\[[0-9.]* [0-9.]*\] 0 d$' shapes-qdf.pdf || fail "the dashed outline is missing"
if grep -E "warning xlsx\.drawing-approximated" "$log"; then fail "the supported shapes should not be reported as approximated"; fi

# Charts: six drawn on one page (titles, legend entries, axis labels, pie percentages, series colors), the scatter chart is the only one reported.
chart_pages=0; chart_page=0
for p in $(seq 1 "$pages"); do [ "$(header_of "$p")" = "Charts" ] && { chart_pages=$((chart_pages + 1)); chart_page=$p; }; done
[ "$chart_pages" = 1 ] || fail "the charts sheet should fit one page, found $chart_pages"
chart_text=$(page_text "$chart_page")
for text in "Clustered columns" "Stacked columns" "Bars with labels" "Lines" "Pie shares" "100 % stacked" "100%" "26%" "Units" "North" "Q6"; do
  [[ $chart_text == *"$text"* ]] || fail "missing chart text: $text"
done
[[ $chart_text != *"Scatter"* ]] || fail "the unsupported scatter chart should not be drawn"
qpdf --qdf --object-streams=disable "$pdf" charts-qdf.pdf
for color in "0.2667 0.4471 0.7686" "0.9294 0.4902 0.1922" "0.6471 0.6471 0.6471"; do
  grep -aq "^$color rg$" charts-qdf.pdf || fail "chart series color $color is missing"
done
grep -E "warning xlsx\.drawing" "$log" | grep -q "scatter chart" || fail "the scatter chart should be reported"
if grep -E "warning xlsx\.(chart|drawing-approximated)" "$log"; then fail "the supported charts should not warn"; fi

# Print areas: the three areas of the "Areas" sheet are three separate pages, and nothing outside them is printed.
area_pages=0
for p in $(seq 1 "$pages"); do
  [ "$(header_of "$p")" = "Areas" ] || continue
  area_pages=$((area_pages + 1))
  case "$area_pages" in
    1) [[ $(page_text "$p") == *"End one"* ]] || fail "the first area should print on its own page";;
    2) [[ $(page_text "$p") == *"End two"* ]] || fail "the second area should print on its own page";;
    3) [[ $(page_text "$p") == *"End three"* ]] || fail "the third area should print on its own page";;
  esac
done
[ "$area_pages" = 3 ] || fail "expected three pages for the three print areas, found $area_pages"
if grep -q "never printed" all.txt; then fail "text outside the print areas was printed"; fi
if grep -E "warning xlsx\.print-area" "$log"; then fail "multiple print areas should not warn any more"; fi

# Comments: printed at the end of the sheet as a list in cell order, on their own page after the cells.
comment_pages=()
for p in $(seq 1 "$pages"); do [ "$(header_of "$p")" = "Comments" ] && comment_pages+=("$p"); done
[ "${#comment_pages[@]}" = 2 ] || fail "the comments sheet should be a page of cells and a page of comments, found ${#comment_pages[@]} pages"
cells_page=$(page_text "${comment_pages[0]}"); list_page=$(page_text "${comment_pages[1]}")
[[ $cells_page == *"Opening"* && $cells_page != *"Starting balance"* ]] || fail "the comment text must not be on the page of cells"
for text in "Cell A3 (Ada)" "Starting balance, carried over." "Cell B4 (Grace)" "Check this figure against the ledger." "Cell C6 (Ada)" "deliberately long"; do
  [[ $list_page == *"$text"* ]] || fail "the comment list is missing: $text"
done
[[ ${list_page%%Cell B4*} == *"Cell A3"* && ${list_page%%Cell C6*} == *"Cell B4"* ]] || fail "comments should be listed in cell order"
if grep -E "warning xlsx\.comments" "$log"; then fail "comments should not warn any more"; fi

# Paper sizes: the "Paper" sheet is A3 (842 x 1191 pt), every other sheet is A4 or Letter.
for p in $(seq 1 "$pages"); do
  if [ "$(header_of "$p")" = "Paper" ]; then
    pdfinfo -f "$p" -l "$p" "$pdf" | grep -E "^Page +$p size:" | grep -qE "(841|842)\.[0-9]+ x (1190|1191)" || fail "the Paper sheet should be A3"
  fi
done
