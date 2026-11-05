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
   - Conditional formatting: cellIs, expression (same-sheet references, arithmetic/comparison/`&`, and AND/OR/NOT/IF/MOD/ROW/COLUMN/ISBLANK/COUNTIF/SUM/… — a fixed function list), top10, aboveAverage (with standard deviations), duplicate/unique values, text/blank/error rules, `timePeriod` rules (relative to `XlsxReaderOptions.Today` or the current date), 2- and 3-color scales, data bars (solid or gradient, borders and negative values with an axis from the x14 extension), every standard 3-, 4- and 5-icon set, and rules that exist only in the x14 extension (3Stars, 3Triangles, 5Boxes, `xm:f` formulas with an inline `x14:dxf`); priority, stopIfTrue and dxf font/fill/border/number-format overrides
   - Shapes and text boxes: preset geometry (rectangle, rounded rectangle, ellipse, triangle, right triangle, diamond, parallelogram, trapezoid, hexagon, the four block arrows, lines and straight connectors), solid fill, outline with width and dash, theme style references, text with alignment/anchor/insets, rotation, flips and groups. Gaps: custom geometry is skipped (`xlsx.drawing`); other preset geometries are drawn as rectangles and bent/curved connectors as straight lines (`xlsx.drawing-approximated`); gradient, pattern and picture fills use the first color; arrowheads, shadows, effects and rotated groups are ignored
   - Charts: column and bar charts (clustered, stacked, 100 % stacked), line charts (markers) and pie charts, with title, legend, axis titles, gridlines, an Excel-like automatic scale, fixed min/max/major unit, number formats, per-series and per-point colors, and value/percent data labels. Data comes from the chart's cache, else from the cells. **Not drawn** (reported by name in `xlsx.drawing`): area, scatter, bubble, radar, stock, surface, doughnut, pie-of-pie/bar-of-pie, all 3-D charts and the newer chart types (waterfall, funnel, …). Combination charts draw only their first chart type and secondary axes are ignored (`xlsx.drawing-approximated`); trendlines, error bars, smoothed lines, log scales, reversed axes, manual layouts and chart text formatting beyond the base size are ignored
   - Comments: listed after the sheet (cell, author, text) when the page setup prints them (`atEnd`; `asDisplayed` falls back to the list with a warning). Threaded comments show their legacy fallback text, and floating notes are not drawn
   - Not yet supported (each emits a warning): pivot tables (cached values print), unlisted paper sizes (Letter is used; the standard Letter/Legal/A/B/envelope sizes are read), and conditional-formatting rules using other sheets, names or unlisted functions — those rules are ignored
   - Conditional formatting gaps: only cells in the used range or print areas are formatted (so a fill on cells beyond them is not drawn); custom icons, data-bar axis options other than automatic and x14 data bars without a main-namespace rule are ignored; a text cell holding `""` counts as blank; ties and rounding follow Excel only approximately
   - Number formats: locale-specific built-in ids 27–36/50–58 (General), `*` fill ignored, fractions ignore thousands separators
   - Layout: text trimmed per character at cell edges, fill alignment treated as left, print titles outside the print range not repeated, header/footer pictures ignored and headers not scaled, `&D`/`&T` use the current date/time, row autofit is approximate (rotated and rich text included), `firstPageNumber` ignored, fit-to-page never enlarges and ignores manual breaks (like Excel), GIF/BMP/TIFF/EMF pictures draw a placeholder
5. PPTX → PDF / PNG
6. ODF readers, legacy DOC/XLS

## License

[MIT](LICENSE)
