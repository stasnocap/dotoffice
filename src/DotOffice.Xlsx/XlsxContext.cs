using System.Globalization;
using System.Text;
using System.Xml;
using DotOffice.Core;
using DotOffice.Opc;

namespace DotOffice.Xlsx;

/// <summary>A sheet entry of workbook.xml.</summary>
internal sealed record SheetEntry(string Name, string? RelationshipId, SheetVisibility Visibility);

/// <summary>A defined name of workbook.xml, optionally scoped to one sheet (by position).</summary>
internal sealed record DefinedName(string Name, int? LocalSheetId, string Formula);

/// <summary>Package-wide state while reading one workbook.</summary>
internal sealed class XlsxContext(OpcPackage package, XlsxReaderOptions options, string workbookPart)
{
    private readonly List<string> _sharedStrings = [];
    private List<CellFormat> _formats = [];
    private long _cellCount;
    private bool _richTextWarned;

    public OpcPackage Package { get; } = package;

    public XlsxReaderOptions Options { get; } = options;

    public string WorkbookPart { get; } = workbookPart;

    /// <summary>True when the workbook uses the 1904 date system.</summary>
    public bool Date1904 { get; private set; }

    public Workbook ReadWorkbook()
    {
        var workbook = new Workbook();
        ReadCoreProperties(workbook);

        var sheets = new List<SheetEntry>();
        var names = new List<DefinedName>();
        using (var r = Package.OpenXml(WorkbookPart))
            ReadWorkbookXml(r, sheets, names);

        var shared = Package.FindRelationship(WorkbookPart, "sharedStrings");
        if (shared is not null && Package.Exists(shared.Target))
        {
            using var r = Package.OpenXml(shared.Target);
            ReadSharedStrings(r);
        }

        ReadStyles();

        var rels = Package.GetRelationships(WorkbookPart);
        for (int index = 0; index < sheets.Count; index++)
        {
            var entry = sheets[index];
            if (entry.RelationshipId is null || !rels.TryGetValue(entry.RelationshipId, out var rel) || rel.IsExternal)
            {
                Options.Warnings.Warn("xlsx.missing-sheet", $"Sheet '{entry.Name}' has no worksheet part and was skipped.");
                continue;
            }
            if (rel.Kind != "worksheet")
            {
                Options.Warnings.Warn("xlsx.unsupported-sheet", $"Sheet '{entry.Name}' is a {rel.Kind}, which is not supported, and was skipped.");
                continue;
            }
            if (!Package.Exists(rel.Target))
            {
                Options.Warnings.Warn("xlsx.missing-sheet", $"Worksheet part '{rel.Target}' of sheet '{entry.Name}' is missing.");
                continue;
            }

            var sheet = new Worksheet(entry.Name) { Visibility = entry.Visibility };
            _sheetWarnings.Clear();
            using (var r = Package.OpenXml(rel.Target))
                ReadWorksheet(r, sheet);
            WarnAboutDeferredFeatures(sheet, rel.Target);
            ApplyDefinedNames(sheet, index, names);
            workbook.Sheets.Add(sheet);
        }
        return workbook;
    }

    // ------------------------------------------------------------------ workbook.xml

    private void ReadWorkbookXml(XmlReader r, List<SheetEntry> sheets, List<DefinedName> names)
    {
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType != XmlNodeType.Element)
            {
                r.Read();
                continue;
            }
            switch (r.LocalName)
            {
                case "workbookPr":
                    Date1904 = r.BoolAttr("date1904");
                    r.Read();
                    break;
                case "sheet":
                    var name = r.GetAttribute("name");
                    if (string.IsNullOrEmpty(name)) { r.Read(); break; }
                    var visibility = r.GetAttribute("state") switch
                    {
                        "hidden" => SheetVisibility.Hidden,
                        "veryHidden" => SheetVisibility.VeryHidden,
                        _ => SheetVisibility.Visible,
                    };
                    sheets.Add(new SheetEntry(name, r.RelationshipId(), visibility));
                    r.Read();
                    break;
                case "definedName":
                    var definedName = r.GetAttribute("name");
                    var local = r.IntAttr("localSheetId");
                    if (definedName is null || r.IsEmptyElement) { r.Read(); break; }
                    names.Add(new DefinedName(definedName, local, r.ReadElementContentAsString()));
                    break;
                default:
                    r.Read();
                    break;
            }
        }
    }

    private void ReadCoreProperties(Workbook workbook)
    {
        var rel = Package.FindRelationship("", "core-properties");
        if (rel is null || !Package.Exists(rel.Target)) return;
        try
        {
            using var r = Package.OpenXml(rel.Target);
            r.MoveToContent();
            while (!r.EOF)
            {
                if (r.NodeType == XmlNodeType.Element && !r.IsEmptyElement && r.LocalName == "title")
                    workbook.Title = NullIfEmpty(r.ReadElementContentAsString());
                else if (r.NodeType == XmlNodeType.Element && !r.IsEmptyElement && r.LocalName == "creator")
                    workbook.Author = NullIfEmpty(r.ReadElementContentAsString());
                else
                    r.Read();
            }
        }
        catch (XmlException)
        {
            Options.Warnings.Warn("xlsx.core-properties", "The document properties part is malformed and was ignored.");
        }
    }

    private static string? NullIfEmpty(string s) => s.Length == 0 ? null : s;

    // ------------------------------------------------------------------ styles.xml and theme

    private void ReadStyles()
    {
        var theme = XlsxTheme.Default;
        var themeRel = Package.FindRelationship(WorkbookPart, "theme");
        if (themeRel is not null && Package.Exists(themeRel.Target))
        {
            try
            {
                using var r = Package.OpenXml(themeRel.Target);
                theme = XlsxTheme.Read(r);
            }
            catch (XmlException)
            {
                Options.Warnings.Warn("xlsx.theme", "The theme part is malformed; default theme colors are used.");
            }
        }

        var stylesRel = Package.FindRelationship(WorkbookPart, "styles");
        if (stylesRel is null || !Package.Exists(stylesRel.Target)) return;
        using var styles = Package.OpenXml(stylesRel.Target);
        _formats = XlsxStyles.Read(styles, theme, Options.Warnings);
    }

    private CellFormat FormatOf(string? styleAttribute) =>
        int.TryParse(styleAttribute, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < _formats.Count
            ? _formats[index]
            : CellFormat.Default;

    // ------------------------------------------------------------------ sharedStrings.xml

    private void ReadSharedStrings(XmlReader r)
    {
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "si")
            {
                _sharedStrings.Add(r.ReadRichText(out var rich));
                if (rich) WarnRichText();
            }
            else
            {
                r.Read();
            }
        }
    }

    private void WarnRichText()
    {
        if (_richTextWarned) return;
        _richTextWarned = true;
        Options.Warnings.Warn("xlsx.rich-text", "Rich text formatting inside cells is not supported; the text is shown with the cell's own style.");
    }

    // ------------------------------------------------------------------ worksheet

    private readonly HashSet<string> _sheetWarnings = [];

    private void WarnOnce(Worksheet sheet, string code, string message)
    {
        if (_sheetWarnings.Add(code)) Options.Warnings.Warn(code, $"Sheet '{sheet.Name}' {message}");
    }

    /// <summary>Features that are not rendered yet are reported once per sheet and then ignored.</summary>
    private void WarnAboutDeferredFeatures(Worksheet sheet, string sheetPart)
    {
        foreach (var rel in Package.GetRelationships(sheetPart).Values)
        {
            if (rel.IsExternal) continue;
            switch (rel.Kind)
            {
                case "drawing":
                    WarnOnce(sheet, "xlsx.drawing", "contains charts, images or shapes, which are not rendered.");
                    break;
                case "pivotTable":
                    WarnOnce(sheet, "xlsx.pivot", "contains a pivot table; its cached cell values are printed but it is not refreshed or styled.");
                    break;
                case "comments":
                    WarnOnce(sheet, "xlsx.comments", "contains comments, which are not printed.");
                    break;
            }
        }
        if (sheet.Cells.Values.Any(c => c.Style.Rotation != 0))
            WarnOnce(sheet, "xlsx.rotation", "contains rotated text, which is drawn horizontally.");
    }

    private void ReadWorksheet(XmlReader r, Worksheet sheet)
    {
        bool formulaWarned = false;
        sheet.DefaultColumnWidth = 48; // Excel's 8.43-character default (64 px)
        bool fitToPage = false;
        bool breaksAreRows = true;
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType != XmlNodeType.Element) { r.Read(); continue; }
            switch (r.LocalName)
            {
                case "sheetFormatPr":
                    ReadSheetFormat(r, sheet);
                    r.Read();
                    break;
                case "cols":
                    ReadColumns(r, sheet);
                    break;
                case "mergeCells":
                    ReadMergeCells(r, sheet);
                    break;
                case "sheetData":
                    ReadSheetData(r, sheet, ref formulaWarned);
                    break;
                case "drawing":
                    WarnOnce(sheet, "xlsx.drawing", "contains charts, images or shapes, which are not rendered.");
                    r.Read();
                    break;
                case "conditionalFormatting":
                    WarnOnce(sheet, "xlsx.conditional-formatting", "uses conditional formatting, which is not applied.");
                    r.Skip();
                    break;
                case "pageSetUpPr":
                    fitToPage = r.BoolAttr("fitToPage");
                    r.Read();
                    break;
                case "printOptions":
                    sheet.PageSetup.CenterHorizontally = r.BoolAttr("horizontalCentered");
                    sheet.PageSetup.CenterVertically = r.BoolAttr("verticalCentered");
                    sheet.PageSetup.PrintGridLines = r.BoolAttr("gridLines");
                    r.Read();
                    break;
                case "pageMargins":
                    ReadMargins(r, sheet.PageSetup);
                    r.Read();
                    break;
                case "pageSetup":
                    ReadPageSetup(r, sheet, fitToPage);
                    r.Read();
                    break;
                case "headerFooter":
                    ReadHeaderFooter(r, sheet.PageSetup);
                    break;
                case "rowBreaks":
                case "colBreaks":
                    breaksAreRows = r.LocalName == "rowBreaks";
                    r.Read();
                    break;
                case "brk":
                    if (r.IntAttr("id") is { } breakId && breakId >= 1)
                        (breaksAreRows ? sheet.PageSetup.RowBreaks : sheet.PageSetup.ColumnBreaks).Add(breakId - 1);
                    r.Read();
                    break;
                default:
                    r.Read();
                    break;
            }
        }
    }

    private static void ReadMargins(XmlReader r, SheetPageSetup setup)
    {
        double Inches(string name, double fallback) => (r.DoubleAttr(name) is { } v && v >= 0 ? v : fallback) * 72;
        setup.Margins = new Thickness(Inches("top", 0.75), Inches("right", 0.7), Inches("bottom", 0.75), Inches("left", 0.7));
        setup.HeaderDistance = Inches("header", 0.3);
        setup.FooterDistance = Inches("footer", 0.3);
    }

    private void ReadPageSetup(XmlReader r, Worksheet sheet, bool fitToPage)
    {
        var setup = sheet.PageSetup;
        double width, height;
        switch (r.IntAttr("paperSize") ?? 1)
        {
            case 1: (width, height) = (612, 792); break;
            case 3: (width, height) = (792, 1224); break;
            case 5: (width, height) = (612, 1008); break;
            case 7: (width, height) = (522, 756); break;
            case 8: (width, height) = (841.89, 1190.55); break;
            case 9: (width, height) = (595.28, 841.89); break;
            case 11: (width, height) = (419.53, 595.28); break;
            default:
                Options.Warnings.Warn("xlsx.paper-size", $"Sheet '{sheet.Name}' uses an unsupported paper size; Letter is used instead.");
                (width, height) = (612, 792);
                break;
        }
        bool landscape = r.GetAttribute("orientation") == "landscape";
        setup.Orientation = landscape ? Orientation.Landscape : Orientation.Portrait;
        // PageWidth/PageHeight are the effective (oriented) page size.
        (setup.PageWidth, setup.PageHeight) = landscape ? (height, width) : (width, height);
        if (r.IntAttr("scale") is { } scale) setup.Scale = Math.Clamp(scale, 10, 400) / 100.0;
        if (fitToPage)
        {
            setup.FitToWidth = Math.Max(r.IntAttr("fitToWidth") ?? 1, 0);
            setup.FitToHeight = Math.Max(r.IntAttr("fitToHeight") ?? 1, 0);
        }
        setup.PageOrderOverThenDown = r.GetAttribute("pageOrder") == "overThenDown";
    }

    private static void ReadHeaderFooter(XmlReader r, SheetPageSetup setup)
    {
        bool different = r.BoolAttr("differentFirst");
        bool oddEven = r.BoolAttr("differentOddEven");
        var texts = new Dictionary<string, string>();
        if (r.IsEmptyElement) { r.Read(); return; }
        int depth = r.Depth;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName is "oddHeader" or "oddFooter" or "evenHeader" or "evenFooter" or "firstHeader" or "firstFooter")
            {
                string name = r.LocalName;
                if (r.IsEmptyElement) { texts[name] = ""; r.Read(); }
                else texts[name] = r.ReadElementContentAsString();
            }
            else r.Read();
        }
        r.Read();

        HeaderFooterParts Parts(string key) => SplitHeaderFooter(texts.GetValueOrDefault(key, ""));
        setup.Header = new HeaderFooterText
        {
            Odd = Parts("oddHeader"),
            First = different ? Parts("firstHeader") : null,
            Even = oddEven ? Parts("evenHeader") : null,
        };
        setup.Footer = new HeaderFooterText
        {
            Odd = Parts("oddFooter"),
            First = different ? Parts("firstFooter") : null,
            Even = oddEven ? Parts("evenFooter") : null,
        };
    }

    /// <summary>Splits header/footer text at <c>&amp;L</c>, <c>&amp;C</c> and <c>&amp;R</c>; other codes stay in the strings. Text before any section code is centered.</summary>
    internal static HeaderFooterParts SplitHeaderFooter(string text)
    {
        var parts = new[] { new StringBuilder(), new StringBuilder(), new StringBuilder() };
        int current = 1;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c != '&' || i + 1 >= text.Length) { parts[current].Append(c); continue; }
            char next = text[i + 1];
            switch (next)
            {
                case 'L': current = 0; i++; break;
                case 'C': current = 1; i++; break;
                case 'R': current = 2; i++; break;
                case '"':
                {
                    int end = text.IndexOf('"', i + 2);
                    if (end < 0) end = text.Length - 1;
                    parts[current].Append(text, i, end - i + 1);
                    i = end;
                    break;
                }
                default:
                    parts[current].Append(c).Append(next);
                    i++;
                    break;
            }
        }
        return new HeaderFooterParts(parts[0].ToString(), parts[1].ToString(), parts[2].ToString());
    }

    private const double MaxDigitWidth = 7;

    /// <summary>Converts a width in characters of the default font's digit to points (ECMA-376 18.3.1.13).</summary>
    internal static double WidthToPoints(double characters)
    {
        double pixels = Math.Truncate((256 * characters + Math.Truncate(128 / MaxDigitWidth)) / 256 * MaxDigitWidth);
        return pixels * 0.75;
    }

    private static double? DoubleAttr(XmlReader r, string name) =>
        double.TryParse(r.GetAttribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) && v >= 0 ? v : null;

    private static void ReadSheetFormat(XmlReader r, Worksheet sheet)
    {
        if (DoubleAttr(r, "defaultColWidth") is { } width) sheet.DefaultColumnWidth = WidthToPoints(width);
        else if (DoubleAttr(r, "baseColWidth") is { } baseWidth)
            sheet.DefaultColumnWidth = Math.Ceiling((Math.Truncate(baseWidth * MaxDigitWidth) + 5) / 8) * 8 * 0.75;
        if (DoubleAttr(r, "defaultRowHeight") is { } height && height > 0) sheet.DefaultRowHeight = Math.Min(height, 409.5);
    }

    private void ReadColumns(XmlReader r, Worksheet sheet)
    {
        if (r.IsEmptyElement) { r.Read(); return; }
        int depth = r.Depth;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "col")
            {
                int min = (r.IntAttr("min") ?? 1) - 1;
                int max = Math.Min((r.IntAttr("max") ?? min + 1) - 1, CellReference.MaxColumns - 1);
                double? width = DoubleAttr(r, "width");
                bool hidden = r.BoolAttr("hidden");
                for (int c = Math.Max(min, 0); c <= max; c++)
                {
                    if (width is { } w) sheet.ColumnWidths[c] = Math.Min(WidthToPoints(w), 1790);
                    if (hidden) sheet.HiddenColumns.Add(c);
                }
            }
            r.Read();
        }
        r.Read();
    }

    private void ReadMergeCells(XmlReader r, Worksheet sheet)
    {
        if (r.IsEmptyElement) { r.Read(); return; }
        int depth = r.Depth;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "mergeCell")
            {
                if (r.GetAttribute("ref") is { } reference && CellReference.TryParseRange(reference, out var range))
                {
                    if (range.FirstRow != range.LastRow || range.FirstColumn != range.LastColumn) sheet.Merges.Add(range);
                }
                else Options.Warnings.Warn("xlsx.bad-reference", $"Sheet '{sheet.Name}' has an invalid merged range.");
            }
            r.Read();
        }
        r.Read();
    }

    private void ReadSheetData(XmlReader r, Worksheet sheet, ref bool formulaWarned)
    {
        if (r.IsEmptyElement)
        {
            r.Read();
            return;
        }
        int depth = r.Depth;
        int rowIndex = -1;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "row")
            {
                rowIndex = r.IntAttr("r") is { } rowNumber ? rowNumber - 1 : rowIndex + 1;
                if (rowIndex is < 0 or >= CellReference.MaxRows)
                {
                    Options.Warnings.Warn("xlsx.bad-reference", $"Sheet '{sheet.Name}' has a row outside the supported range.");
                    r.Skip();
                    continue;
                }
                if (DoubleAttr(r, "ht") is { } rowHeight) sheet.RowHeights[rowIndex] = Math.Min(rowHeight, 409.5);
                if (r.BoolAttr("hidden")) sheet.HiddenRows.Add(rowIndex);
                ReadRow(r, sheet, rowIndex, ref formulaWarned);
            }
            else
            {
                r.Read();
            }
        }
        r.Read();
    }

    /// <summary>Reads the cells of a <c>row</c>; leaves the reader after the row's end tag.</summary>
    private void ReadRow(XmlReader r, Worksheet sheet, int rowIndex, ref bool formulaWarned)
    {
        if (r.IsEmptyElement)
        {
            r.Read();
            return;
        }
        int depth = r.Depth;
        int column = -1;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType != XmlNodeType.Element || r.LocalName != "c")
            {
                r.Read();
                continue;
            }

            int row = rowIndex;
            int col = column + 1;
            var reference = r.GetAttribute("r");
            if (reference is not null)
            {
                if (CellReference.TryParse(reference, out var refRow, out var refCol)) (row, col) = (refRow, refCol);
                else Options.Warnings.Warn("xlsx.bad-reference", $"Sheet '{sheet.Name}' has an invalid cell reference '{reference}'.");
            }
            column = col;

            var cell = ReadCell(r, sheet, ref formulaWarned);
            if (cell is null || col >= CellReference.MaxColumns) continue;
            if (++_cellCount > Options.MaxCellCount)
                throw new InvalidDataException($"The workbook has more than the limit of {Options.MaxCellCount} cells.");
            sheet.Cells[(row, col)] = cell;
        }
        r.Read();
    }

    /// <summary>Reads one <c>c</c> element, leaving the reader after its end tag; null for cells without a displayable value.</summary>
    private Cell? ReadCell(XmlReader r, Worksheet sheet, ref bool formulaWarned)
    {
        var type = r.GetAttribute("t") ?? "n";
        var format = FormatOf(r.GetAttribute("s"));
        if (r.IsEmptyElement)
        {
            r.Read();
            return BuildEmpty(format);
        }

        int depth = r.Depth;
        string? value = null;
        string? inline = null;
        bool hasFormula = false;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType != XmlNodeType.Element)
            {
                r.Read();
                continue;
            }
            switch (r.LocalName)
            {
                case "v":
                    if (r.IsEmptyElement) { value = ""; r.Read(); }
                    else value = r.ReadElementContentAsString();
                    break;
                case "f":
                    hasFormula = true;
                    r.Skip();
                    break;
                case "is":
                    inline = r.ReadRichText(out var rich);
                    if (rich) WarnRichText();
                    break;
                default:
                    r.Skip();
                    break;
            }
        }
        r.Read();
        return Build(type, value, inline, hasFormula, format, sheet, ref formulaWarned);
    }

    /// <summary>An empty cell only matters when it paints something (fill or border).</summary>
    private static Cell? BuildEmpty(CellFormat format) =>
        format.Style.Fill is not null || format.Style.Borders.HasAny ? new Cell { Style = format.Style } : null;

    private static string CodeOf(CellFormat format) => format.NumberFormatCode ?? NumberFormat.BuiltIn(format.NumberFormatId) ?? "General";

    private static CellStyle WithColor(CellStyle style, Color? color) => color is { } c ? style with { Font = style.Font with { Color = c } } : style;

    private static Cell TextCell(string text, CellFormat format)
    {
        var code = CodeOf(format);
        if (code == "General") return new Cell { Text = text, Kind = CellValueKind.Text, Style = format.Style };
        var shown = NumberFormat.Parse(code).FormatText(text);
        return new Cell { Text = shown.Text, Kind = CellValueKind.Text, Style = WithColor(format.Style, shown.Color) };
    }

    private Cell? Build(string type, string? value, string? inline, bool hasFormula, CellFormat format, Worksheet sheet, ref bool formulaWarned)
    {
        var style = format.Style;
        if (type == "inlineStr")
            return inline is null ? BuildEmpty(format) : TextCell(inline, format);
        if (value is null)
        {
            if (hasFormula && !formulaWarned)
            {
                formulaWarned = true;
                Options.Warnings.Warn("xlsx.formula-no-cache", $"Sheet '{sheet.Name}' has formulas without a cached value; those cells are shown empty.");
            }
            return BuildEmpty(format);
        }

        switch (type)
        {
            case "s":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < _sharedStrings.Count)
                    return TextCell(_sharedStrings[index], format);
                Options.Warnings.Warn("xlsx.bad-shared-string", $"Sheet '{sheet.Name}' refers to a shared string that does not exist.");
                return BuildEmpty(format);
            case "str":
            case "d":
                return TextCell(XlsxXml.DecodeEscapes(value), format);
            case "b":
                return new Cell { Text = value.Trim() is "1" or "true" ? "TRUE" : "FALSE", Kind = CellValueKind.Bool, Style = style };
            case "e":
                return new Cell { Text = value, Kind = CellValueKind.Error, Style = style };
            default:
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
                {
                    var shown = NumberFormat.Parse(CodeOf(format)).FormatNumber(number, Date1904);
                    return new Cell { Text = shown.Text, Kind = CellValueKind.Number, Style = WithColor(style, shown.Color) };
                }
                return value.Length == 0 ? BuildEmpty(format) : new Cell { Text = value, Kind = CellValueKind.Text, Style = style };
        }
    }

    // ------------------------------------------------------------------ defined names

    private void ApplyDefinedNames(Worksheet sheet, int sheetPosition, List<DefinedName> names)
    {
        foreach (var name in names)
        {
            if (name.LocalSheetId != sheetPosition) continue;
            if (name.Name == "_xlnm.Print_Area")
            {
                var areas = CellReference.ParseAreas(name.Formula).Where(a => a.Range is not null).ToList();
                if (areas.Count == 0) continue;
                if (areas.Count > 1)
                    Options.Warnings.Warn("xlsx.print-area-multiple", $"Sheet '{sheet.Name}' has several print areas; only the first is used.");
                sheet.PageSetup.PrintArea = areas[0].Range;
            }
            else if (name.Name == "_xlnm.Print_Titles")
            {
                foreach (var area in CellReference.ParseAreas(name.Formula))
                {
                    if (area.RowSpan is { } rows) sheet.PageSetup.TitleRows = rows;
                    else if (area.ColumnSpan is { } columns) sheet.PageSetup.TitleColumns = columns;
                }
            }
        }
    }
}
