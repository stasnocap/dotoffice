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
3. DOCX polish: fields, floats, columns, footnotes, complex scripts ✓ (Indic shaping not yet supported)
4. XLSX → PDF
5. PPTX → PDF / PNG
6. ODF readers, legacy DOC/XLS

## License

[MIT](LICENSE)
