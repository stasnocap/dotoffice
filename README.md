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
   - Conditional formatting: cellIs, expression (same-sheet references, arithmetic/comparison/`&`, and AND/OR/NOT/IF/MOD/ROW/COLUMN/ISBLANK/COUNTIF/SUM/… — a fixed function list), top10, aboveAverage (with standard deviations), duplicate/unique values, text/blank/error rules, 2- and 3-color scales, data bars (solid or gradient from the x14 extension, negative values with an axis) and the 3Arrows, 3TrafficLights1 and 3Symbols icon sets; priority, stopIfTrue and dxf font/fill/border/number-format overrides
   - Not yet supported (each emits a warning): charts and shapes (`xlsx.drawing`), pivot tables (cached values print), comments, flipped pictures (drawn unflipped), multiple print areas (only the first), unlisted paper sizes (Letter; the standard Letter/Legal/A/B/envelope sizes are read), and conditional-formatting rules of other kinds (`timePeriod`, other icon sets, formulas using other sheets, names or unlisted functions) — those rules are ignored
   - Conditional formatting gaps: only cells in the used range or print area are formatted (so a fill on cells beyond both is not drawn); custom icons, data-bar borders and axis options other than automatic are ignored; extension-only rules (x14 icon sets, 4+ icon sets) are not read; a text cell holding `""` counts as blank; ties and rounding follow Excel only approximately
   - Number formats: locale-specific built-in ids 27–36/50–58 (General), `*` fill ignored, `[ColorN]` uses the default palette (a workbook's custom palette is not consulted), fractions ignore thousands separators
   - Layout: text trimmed per character at cell edges, fill alignment treated as left, print titles outside the print range not repeated, header/footer pictures ignored and headers not scaled, `&D`/`&T` use the current date/time, row autofit is approximate (rotated and rich text included), `firstPageNumber` ignored, fit-to-page never enlarges and ignores manual breaks (like Excel), GIF/BMP/TIFF/EMF pictures draw a placeholder
5. PPTX → PDF / PNG
6. ODF readers, legacy DOC/XLS

## License

[MIT](LICENSE)
