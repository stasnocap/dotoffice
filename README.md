# DotOffice

Pure managed .NET library for converting office documents — no LibreOffice, no native binaries, MIT licensed.

> **Status:** early development. Not usable yet.

```csharp
await DocumentConverter.ConvertAsync("invoice.docx", "invoice.pdf");
```

## Goals

- DOCX/XLSX/PPTX → PDF with high fidelity
- 100% managed code: runs anywhere .NET runs (containers, Linux, AOT)
- No GPL/AGPL or restrictive-license dependencies

## Packages

| Package | Purpose |
|---|---|
| `DotOffice.Core` | Format-neutral document model |
| `DotOffice.Fonts` | OpenType parsing, subsetting, font resolution |
| `DotOffice.Pdf` | Low-level PDF writer |
| `DotOffice.Layout` | Line breaking, pagination, tables |
| `DotOffice.Docx` / `Xlsx` / `Pptx` | OOXML readers |
| `DotOffice.Converter` | Public facade |

## Roadmap

1. PDF writer + font embedding/subsetting ✓
2. DOCX → PDF: paragraphs, styles, lists, tables, headers/footers, images ✓
3. DOCX polish: fields, floats, columns, footnotes, complex scripts including Indic ✓ (Sinhala, Tibetan, Myanmar and Khmer shaping not yet supported)
4. XLSX → PDF: milestone 1 ✓ (cells, styles, number formats, merges, page setup, print areas/titles, breaks, headers/footers); milestone 2 ✓ (rich-text runs, rotated and stacked text, gradient fills, pictures from drawings with anchors/crop/rotation, conditional formatting, fraction formats)
   - Milestone 3 ✓ (flipped pictures, multiple print areas, standard paper sizes, the workbook's `[ColorN]` palette, comments as an end-of-sheet list, shapes and charts, and the conditional-formatting leftovers below)
   - Milestone 4 ✓ (chart types and options, shape fills, arrowheads and paths, conditional formatting across sheets, custom icons and data-bar axes, pivot tables without a blanket warning, all paper sizes, first page numbers, fill alignment)
   - Conditional formatting: cellIs, expression (references on this or another sheet including quoted sheet names, defined names that stand for a reference, a range or a constant, arithmetic/comparison/`&`, and AND/OR/NOT/IF/MOD/ROW/COLUMN/ISBLANK/COUNTIF/SUM/… — a fixed function list; all rules are applied once every sheet has been read), top10, aboveAverage (with standard deviations), duplicate/unique values, text/blank/error rules, `timePeriod` rules (relative to `XlsxReaderOptions.Today` or the current date), 2- and 3-color scales, data bars (solid or gradient, borders, negative values, axis position automatic/middle/none, axis color and right-to-left direction from the x14 extension), every standard 3-, 4- and 5-icon set including custom icons (`x14:cfIcon`, also "no icon"), and rules that exist only in the x14 extension (3Stars, 3Triangles, 5Boxes, `xm:f` formulas with an inline `x14:dxf`); priority, stopIfTrue and dxf font/fill/border/number-format overrides
   - Shapes and text boxes: preset geometry (rectangle, rounded rectangle, ellipse, triangle, right triangle, diamond, parallelogram, trapezoid, hexagon, the four block arrows, lines and connectors), custom geometry (paths of lines, curves and arcs with plain numbers), bent connectors 2–5 and curved connectors 2–3 with their adjustments, solid fills and real gradient fills (linear; path gradients are drawn radially, clipped to the outline), outline with width and dash, arrowheads (triangle, stealth, arrow, oval, diamond; the line stops at the base of a filled head), theme style references, text with alignment/anchor/insets, rotation, flips and groups. Gaps: custom geometry that uses guide formulas is skipped (`xlsx.drawing`); other preset geometries are drawn as rectangles, curved connectors 4–5 are rounded bent routes, rectangular path gradients are radial (`xlsx.drawing-approximated`); pattern and picture fills use the first color, gradient fills from theme style references are solid; custom-geometry paths with holes are filled without them; shadows, effects and rotated groups are ignored
   - Charts: column and bar charts (clustered, stacked, 100 % stacked), line charts (markers, smoothed lines), area, pie, doughnut (rings per series, hole size, first-slice angle), scatter (marker-only or joined, own horizontal scale), bubble (area follows the size) and radar charts (open or filled), and combination charts of bar, line and area groups with a secondary value axis (own scale, labels, title, number format); title, legend, axis titles, gridlines, an Excel-like automatic scale, fixed min/max/major unit, number formats, reversed and logarithmic axes, per-series and per-point colors, and value/percent data labels. Data comes from the chart's cache, else from the cells. **Not drawn** (reported by name in `xlsx.drawing`): stock, surface, pie-of-pie/bar-of-pie, all 3-D charts and the newer chart types (waterfall, funnel, …); a combination with a group of another kind (say scatter) draws only the category-axis groups (`xlsx.drawing-approximated`). Approximations: with a reversed value axis the category axis stays at the bottom, smoothing is a Catmull-Rom spline, log axes are not combined with stacking; trendlines, error bars, manual layouts, transparency, bubble 3-D, secondary axes of horizontal bars and chart text formatting beyond the base size are ignored
   - Comments: listed after the sheet (cell, author, text) when the page setup prints them (`atEnd`; `asDisplayed` falls back to the list with a warning). Threaded comments show their legacy fallback text, and floating notes are not drawn
   - Not yet supported (each emits a warning): conditional-formatting rules using unlisted functions or names that are not plain references (those rules are ignored) and rules reading a sheet that does not exist; pivot tables print their cached cells without a warning, and warn only when the pivot style or pivot-area formats are not applied or the cache refreshes on load (the table's own look is then not reproduced); paper size ids outside the standard list use Letter
   - Paper sizes: every SpreadsheetML `paperSize` id with a fixed size (1–47, 50–118); the "rotated" ids take their upright twin's size and the page's own orientation decides
   - Conditional formatting gaps: only cells in the used range or print areas are formatted (so a fill on cells beyond them is not drawn); x14 data bars without a main-namespace rule are ignored; a text cell holding `""` counts as blank; ties and rounding follow Excel only approximately
   - Number formats: locale-specific built-in ids 27–36/50–58 (General), `*` fill ignored, fractions ignore thousands separators
   - Layout: text trimmed per character at cell edges, print titles outside the print range not repeated, header/footer pictures ignored and headers not scaled, `&D`/`&T` use the current date/time, row autofit is approximate (rotated and rich text included), text with fill alignment repeats only without rich-text runs, fit-to-page never enlarges and ignores manual breaks (like Excel), GIF/BMP/TIFF/EMF pictures draw a placeholder
5. PPTX → PDF / PNG
6. ODF readers, legacy DOC/XLS

## License

[MIT](LICENSE)
