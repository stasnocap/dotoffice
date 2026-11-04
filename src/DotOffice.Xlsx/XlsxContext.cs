using System.Globalization;
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
            using (var r = Package.OpenXml(rel.Target))
                ReadWorksheet(r, sheet);
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

    private void ReadWorksheet(XmlReader r, Worksheet sheet)
    {
        bool formulaWarned = false;
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "sheetData") ReadSheetData(r, sheet, ref formulaWarned);
            else r.Read();
        }
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

    private Cell? Build(string type, string? value, string? inline, bool hasFormula, CellFormat format, Worksheet sheet, ref bool formulaWarned)
    {
        var style = format.Style;
        if (type == "inlineStr")
            return inline is null ? BuildEmpty(format) : new Cell { Text = inline, Kind = CellValueKind.Text, Style = style };
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
                    return new Cell { Text = _sharedStrings[index], Kind = CellValueKind.Text, Style = style };
                Options.Warnings.Warn("xlsx.bad-shared-string", $"Sheet '{sheet.Name}' refers to a shared string that does not exist.");
                return BuildEmpty(format);
            case "str":
            case "d":
                return new Cell { Text = XlsxXml.DecodeEscapes(value), Kind = CellValueKind.Text, Style = style };
            case "b":
                return new Cell { Text = value.Trim() is "1" or "true" ? "TRUE" : "FALSE", Kind = CellValueKind.Bool, Style = style };
            case "e":
                return new Cell { Text = value, Kind = CellValueKind.Error, Style = style };
            default:
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
                    return new Cell { Text = NumberFormatter.General(number), Kind = CellValueKind.Number, Style = style };
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
