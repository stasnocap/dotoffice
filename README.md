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
4. XLSX → PDF: milestone 1 ✓ (cells, styles, number formats, merges, page setup, print areas/titles, breaks, headers/footers); milestone 2 in progress: rich-text runs ✓, rotated and stacked text ✓, gradient fills ✓, pictures from drawings ✓ (PNG/JPEG, anchors, crop, rotation); conditional formatting and fraction formats still to do
   - Not yet supported (each emits a warning): charts and shapes (`xlsx.drawing`), conditional formatting, pivot tables (cached values print), comments, flipped pictures (drawn unflipped), multiple print areas (only the first), unsupported paper sizes (Letter)
   - Number formats: fractions, indexed `[ColorN]`, locale-specific built-in ids 27–36/50–58 (General), `*` fill ignored
   - Layout: text trimmed per character at cell edges, fill alignment treated as left, print titles outside the print range not repeated, header/footer pictures ignored and headers not scaled, `&D`/`&T` use the current date/time, row autofit is approximate (rotated and rich text included), `firstPageNumber` ignored, fit-to-page never enlarges and ignores manual breaks (like Excel), GIF/BMP/TIFF/EMF pictures draw a placeholder
5. PPTX → PDF / PNG
6. ODF readers, legacy DOC/XLS

## License

[MIT](LICENSE)
