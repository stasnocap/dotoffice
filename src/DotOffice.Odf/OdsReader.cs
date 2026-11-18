using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Opc;
using DotOffice.Xlsx;

namespace DotOffice.Odf;

/// <summary>Reads OpenDocument spreadsheets (.ods, .ots) into the <see cref="Workbook"/> model, using the cached cell values.</summary>
public static class OdsReader
{
    /// <summary>Reads an ODS package. Malformed, encrypted or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Workbook Read(Stream stream, OdfReaderOptions? options = null)
    {
        options ??= new OdfReaderOptions();
        try
        {
            using var package = OdfPackage.Open(stream, options, "application/vnd.oasis.opendocument.spreadsheet", out var styles);
            using var content = package.OpenXml("content.xml");
            return new OdsContext(styles, options).Read(content);
        }
        catch (XmlException e)
        {
            throw new InvalidDataException($"The spreadsheet contains malformed XML: {e.Message}", e);
        }
    }
}

internal sealed class OdsContext(OdfStyles styles, OdfReaderOptions options)
{
    private const int MaxRows = 1_048_576, MaxColumns = 16_384, MaxStyledEmpty = 200_000, MaxSheets = 1000, MaxCellText = 32_767;
    private readonly Dictionary<string, CellStyle> _cellStyles = [];
    private long _cells, _styledEmpty;
    private bool _drawingWarned;

    private static readonly RunProperties DefaultFont = new() { FontFamily = "Liberation Sans", FontSize = 10 };

    public Workbook Read(XmlReader reader)
    {
        var workbook = new Workbook();
        var tables = new List<(string Name, string? Style)>();
        reader.MoveToContent();
        if (reader.LocalName != "document-content") throw new InvalidDataException("content.xml is not an OpenDocument content part.");
        int depth = reader.Depth;
        reader.Read();
        while (!reader.EOF && reader.Depth > depth)
        {
            if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
            if (reader.LocalName is "automatic-styles" or "font-face-decls") styles.Add((XElement)XNode.ReadFrom(reader));
            else if (reader.LocalName == "body") ReadBody(reader, workbook);
            else reader.Skip();
        }
        if (workbook.Sheets.Count == 0) throw new InvalidDataException("The document has no spreadsheet tables.");
        return workbook;
    }

    private void ReadBody(XmlReader reader, Workbook workbook)
    {
        int depth = reader.Depth;
        bool found = false;
        if (reader.IsEmptyElement) { reader.Read(); return; }
        reader.Read();
        while (!reader.EOF && reader.Depth > depth)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "spreadsheet") { found = true; ReadSpreadsheet(reader, workbook); }
            else if (reader.NodeType == XmlNodeType.Element) { reader.Skip(); }
            else reader.Read();
        }
        if (!found) throw new InvalidDataException("The document is not a spreadsheet.");
    }

    private void ReadSpreadsheet(XmlReader reader, Workbook workbook)
    {
        int depth = reader.Depth;
        if (reader.IsEmptyElement) { reader.Read(); return; }
        reader.Read();
        while (!reader.EOF && reader.Depth > depth)
        {
            if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
            if (reader.LocalName == "table" && reader.NamespaceURI == Ns.Table.NamespaceName)
            {
                if (workbook.Sheets.Count >= MaxSheets) throw new InvalidDataException("Too many sheets.");
                workbook.Sheets.Add(ReadTable(reader));
            }
            else reader.Skip();
        }
    }

    private Worksheet ReadTable(XmlReader reader)
    {
        var name = reader.GetAttribute("name", Ns.Table.NamespaceName) ?? "Sheet";
        var tableStyle = styles.Properties("table", reader.GetAttribute("style-name", Ns.Table.NamespaceName));
        var sheet = new Worksheet(name);
        if (tableStyle.TryGetValue("table-properties/display", out var display) && display == "false") sheet.Visibility = SheetVisibility.Hidden;
        ApplyPageSetup(sheet, tableStyle.GetValueOrDefault("table-properties/master-page-name") ?? styles.DefaultMasterPage);
        sheet.DefaultColumnWidth = 64.26;
        sheet.DefaultRowHeight = 12.8;

        var columns = new List<(int Column, int Count, string? CellStyle)>();
        int row = 0, column = 0;
        int depth = reader.Depth;
        if (reader.IsEmptyElement) { reader.Read(); return sheet; }
        reader.Read();
        ReadTableChildren(reader, depth, sheet, columns, ref row, ref column);
        return sheet;
    }

    /// <summary>Reads rows and columns, descending into header, group and row-group wrappers.</summary>
    private void ReadTableChildren(XmlReader reader, int depth, Worksheet sheet, List<(int, int, string?)> columns, ref int row, ref int column)
    {
        while (!reader.EOF && reader.Depth > depth)
        {
            if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
            switch (reader.LocalName)
            {
                case "table-column":
                    ReadColumn(reader, sheet, columns, ref column);
                    break;
                case "table-row":
                    ReadRow(reader, sheet, columns, ref row);
                    break;
                case "table-header-columns" or "table-columns" or "table-column-group" or "table-header-rows" or "table-rows" or "table-row-group":
                {
                    int d = reader.Depth;
                    if (reader.IsEmptyElement) { reader.Read(); break; }
                    reader.Read();
                    ReadTableChildren(reader, d, sheet, columns, ref row, ref column);
                    break;
                }
                case "shapes" or "forms":
                    reader.Skip();
                    break;
                default:
                    reader.Skip();
                    break;
            }
        }
        if (reader.NodeType == XmlNodeType.EndElement) reader.Read();
    }

    private void ReadColumn(XmlReader reader, Worksheet sheet, List<(int, int, string?)> columns, ref int column)
    {
        int count = Math.Clamp(Repeat(reader, "number-columns-repeated"), 1, MaxColumns);
        var style = styles.Properties("table-column", reader.GetAttribute("style-name", Ns.Table.NamespaceName));
        var defaultCell = reader.GetAttribute("default-cell-style-name", Ns.Table.NamespaceName);
        var visibility = reader.GetAttribute("visibility", Ns.Table.NamespaceName);
        double? width = OdfUnits.Length(style.GetValueOrDefault("table-column-properties/column-width"));
        for (int i = 0; i < count && column + i < MaxColumns; i++)
        {
            if (width is { } w) sheet.ColumnWidths[column + i] = w;
            if (visibility is "collapse" or "filter") sheet.HiddenColumns.Add(column + i);
        }
        columns.Add((column, count, defaultCell));
        column = Math.Min(column + count, MaxColumns);
        reader.Skip();
    }

    private void ReadRow(XmlReader reader, Worksheet sheet, List<(int Column, int Count, string? CellStyle)> columns, ref int row)
    {
        int repeat = Math.Clamp(Repeat(reader, "number-rows-repeated"), 1, MaxRows);
        var style = styles.Properties("table-row", reader.GetAttribute("style-name", Ns.Table.NamespaceName));
        var rowDefault = reader.GetAttribute("default-cell-style-name", Ns.Table.NamespaceName);
        var visibility = reader.GetAttribute("visibility", Ns.Table.NamespaceName);
        double? height = OdfUnits.Length(style.GetValueOrDefault("table-row-properties/row-height"))
                         ?? OdfUnits.Length(style.GetValueOrDefault("table-row-properties/min-row-height"));
        var element = reader.IsEmptyElement ? null : (XElement)XNode.ReadFrom(reader);
        if (element is null) reader.Read();

        // Rows past the grid are dropped; a repeated blank row only moves the counter.
        if (row >= MaxRows) return;
        int last = Math.Min(row + repeat, MaxRows);
        if (element is not null) ReadCells(element, sheet, columns, rowDefault, row, repeat);
        // Row sizes apply to every repeated row (blank repeats too, up to a cap, since they take up space).
        if (height is not null || visibility is not null)
            for (int r = row; r < last && r - row < 255; r++)
            {
                if (height is { } h) sheet.RowHeights[r] = h;
                if (visibility is "collapse" or "filter") sheet.HiddenRows.Add(r);
            }
        row = last;
    }

    private void ReadCells(XElement rowElement, Worksheet sheet, List<(int Column, int Count, string? CellStyle)> columns, string? rowDefault,
        int row, int rowRepeat)
    {
        int column = 0;
        foreach (var e in rowElement.Elements())
        {
            bool covered = e.Name == Ns.Table + "covered-table-cell";
            if (!covered && e.Name != Ns.Table + "table-cell") continue;
            int repeat = Math.Clamp(RepeatOf(e, "number-columns-repeated"), 1, MaxColumns);
            if (column >= MaxColumns) break;
            if (covered) { column += repeat; continue; }

            var styleName = (string?)e.Attribute(Ns.Table + "style-name") ?? rowDefault ?? DefaultCellStyleOf(columns, column);
            var style = CellStyleOf(styleName);
            var cell = BuildCell(e, style, styleName);
            int colSpan = Math.Max(1, RepeatOf(e, "number-columns-spanned")), rowSpan = Math.Max(1, RepeatOf(e, "number-rows-spanned"));
            bool hasContent = cell is not null && (cell.Kind != CellValueKind.Text || cell.Text.Length > 0 || cell.IsEmptyString);
            if (cell is null) cell = new Cell { Style = style };
            bool styled = style.Fill is not null || style.Borders.HasAny;

            if (hasContent || styled)
            {
                // A blank styled run that spans (nearly) the rest of the row is a column or row default, not formatting someone drew.
                if (!hasContent && (repeat >= 256 || rowRepeat >= 256)) { column += repeat; continue; }
                int copies = repeat;
                int rows = hasContent ? Math.Min(rowRepeat, MaxRows - row) : Math.Min(rowRepeat, 1);
                // A blank styled cell is only kept up to a budget (one spreadsheet may style a whole row of a million columns).
                if (!hasContent && (_styledEmpty += (long)copies * rows) > MaxStyledEmpty) { column += repeat; continue; }
                for (int rr = 0; rr < rows; rr++)
                    for (int c = 0; c < copies && column + c < MaxColumns; c++)
                    {
                        if (hasContent && ++_cells > options.MaxCellCount) throw new InvalidDataException($"The workbook has more than {options.MaxCellCount} cells.");
                        sheet.Cells[(row + rr, column + c)] = cell;
                    }
            }
            if ((colSpan > 1 || rowSpan > 1) && column + colSpan - 1 < MaxColumns)
                sheet.Merges.Add(new CellRange(row, column, Math.Min(row + rowSpan - 1, MaxRows - 1), Math.Min(column + colSpan - 1, MaxColumns - 1)));
            if (e.Elements().Any(x => x.Name.Namespace.NamespaceName.EndsWith(":drawing:1.0", StringComparison.Ordinal)) && !_drawingWarned)
            {
                _drawingWarned = true;
                options.Warnings.Warn("ods.drawing-skipped", $"Sheet '{sheet.Name}' has drawings, images or charts in cells; they are not drawn.");
            }
            column += repeat;
        }
    }

    private static string? DefaultCellStyleOf(List<(int Column, int Count, string? CellStyle)> columns, int column)
    {
        foreach (var (first, count, style) in columns)
            if (column >= first && column < first + count) return style;
        return null;
    }

    private Cell? BuildCell(XElement e, CellStyle style, string? styleName)
    {
        var type = (string?)e.Attribute(Ns.Office + "value-type");
        if (type is null) return null;
        string? dataStyle = null;
        if (styleName is not null && styles.Properties("table-cell", styleName).TryGetValue("style/data-style-name", out var ds)) dataStyle = ds;
        var code = styles.NumberCode(dataStyle) ?? "General";
        var palette = Array.Empty<Color?>();

        switch (type)
        {
            case "float" or "percentage" or "currency":
            {
                if (!double.TryParse((string?)e.Attribute(Ns.Office + "value"), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) || !double.IsFinite(v))
                    return null;
                if (type == "percentage" && code == "General") code = "0%";
                var shown = NumberFormat.Parse(code, palette).FormatNumber(v, false);
                return new Cell { Text = shown.Text, Kind = CellValueKind.Number, Value = v, Style = Colored(style, shown.Color) };
            }
            case "date" or "time":
            {
                double? serial = type == "date" ? DateSerial((string?)e.Attribute(Ns.Office + "date-value")) : TimeSerial((string?)e.Attribute(Ns.Office + "time-value"));
                if (serial is not { } v) return null;
                if (code == "General") code = type == "date" ? "yyyy\\-mm\\-dd" : "hh:mm:ss";
                var shown = NumberFormat.Parse(code, palette).FormatNumber(v, false);
                return new Cell { Text = shown.Text, Kind = CellValueKind.Number, Value = v, Style = Colored(style, shown.Color) };
            }
            case "boolean":
                return new Cell { Text = (string?)e.Attribute(Ns.Office + "boolean-value") == "true" ? "TRUE" : "FALSE", Kind = CellValueKind.Bool, Style = style };
            case "string":
            {
                var text = (string?)e.Attribute(Ns.Office + "string-value") ?? ParagraphText(e);
                if (text.Length > MaxCellText) text = text[..MaxCellText];
                return new Cell { Text = text, Kind = CellValueKind.Text, Style = style, IsEmptyString = text.Length == 0 };
            }
            default:
                return null;
        }
    }

    private static CellStyle Colored(CellStyle style, Color? color) =>
        color is { } c ? style with { Font = style.Font with { Color = c } } : style;

    private static double? DateSerial(string? value)
    {
        if (value is null || !DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var d)) return null;
        return (d - new DateTime(1899, 12, 30)).TotalDays;
    }

    private static double? TimeSerial(string? value)
    {
        if (value is null) return null;
        try { return XmlConvert.ToTimeSpan(value).TotalDays; }
        catch (FormatException) { return null; }
        catch (OverflowException) { return null; }
    }

    /// <summary>Plain text of a cell's paragraphs, joined by line breaks, with the spacing elements expanded.</summary>
    private static string ParagraphText(XElement cell)
    {
        var sb = new StringBuilder();
        bool first = true;
        foreach (var p in cell.Elements().Where(x => x.Name == Ns.Text + "p" || x.Name == Ns.Text + "h"))
        {
            if (!first) sb.Append('\n');
            first = false;
            Append(p, sb);
            if (sb.Length > MaxCellText) break;
        }
        return sb.ToString();
    }

    private static void Append(XElement e, StringBuilder sb)
    {
        foreach (var n in e.Nodes())
        {
            if (sb.Length > MaxCellText) return;
            if (n is XText t) sb.Append(t.Value);
            else if (n is XElement x)
            {
                if (x.Name == Ns.Text + "s") sb.Append(' ', Math.Clamp(RepeatOf(x, "c", Ns.Text), 1, 1000));
                else if (x.Name == Ns.Text + "tab") sb.Append('\t');
                else if (x.Name == Ns.Text + "line-break") sb.Append('\n');
                else if (x.Name.Namespace == Ns.Text) Append(x, sb);
            }
        }
    }

    private CellStyle CellStyleOf(string? name)
    {
        name ??= "";
        if (_cellStyles.TryGetValue(name, out var cached)) return cached;
        var p = styles.Properties("table-cell", name.Length == 0 ? "Default" : name);
        var font = DefaultFont;
        string? family = p.GetValueOrDefault("text-properties/font-family") ?? p.GetValueOrDefault("text-properties/font-name");
        if (family is not null) font = font with { FontFamily = styles.FontFamily(family) };
        if (OdfUnits.Length(p.GetValueOrDefault("text-properties/font-size")) is { } size && size > 0) font = font with { FontSize = size };
        if (p.GetValueOrDefault("text-properties/font-weight") is { } w) font = font with { Bold = w == "bold" || (int.TryParse(w, out var wn) && wn >= 600) };
        if (p.GetValueOrDefault("text-properties/font-style") is { } fs) font = font with { Italic = fs is "italic" or "oblique" };
        if (p.GetValueOrDefault("text-properties/text-underline-style") is { } us) font = font with { Underline = us == "none" ? UnderlineStyle.None : UnderlineStyle.Single };
        if (p.GetValueOrDefault("text-properties/text-line-through-style") is { } ls) font = font with { Strike = ls != "none" };
        if (p.TryGetValue("text-properties/color", out var color)) font = font with { Color = OdfUnits.Hex(color) };

        var borders = new BorderSet
        {
            Top = Edge(p, "top"), Bottom = Edge(p, "bottom"), Left = Edge(p, "left"), Right = Edge(p, "right"),
        };
        var align = p.GetValueOrDefault("paragraph-properties/text-align") switch
        {
            "start" or "left" => HorizontalAlignment.Left,
            "end" or "right" => HorizontalAlignment.Right,
            "center" => HorizontalAlignment.Center,
            "justify" => HorizontalAlignment.Justify,
            _ => HorizontalAlignment.General,
        };
        var style = new CellStyle
        {
            Font = font,
            Fill = OdfUnits.Hex(p.GetValueOrDefault("table-cell-properties/background-color")),
            Borders = borders,
            HorizontalAlignment = align,
            VerticalAlignment = p.GetValueOrDefault("table-cell-properties/vertical-align") switch
            {
                "top" => CellVerticalAlignment.Top,
                "middle" => CellVerticalAlignment.Center,
                _ => CellVerticalAlignment.Bottom,
            },
            Wrap = p.GetValueOrDefault("table-cell-properties/wrap-option") == "wrap",
            Rotation = double.TryParse(p.GetValueOrDefault("table-cell-properties/rotation-angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out var rot) && double.IsFinite(rot) ? rot : 0,
        };
        return _cellStyles[name] = style;
    }

    private static Border Edge(Dictionary<string, string> p, string side)
    {
        var value = p.GetValueOrDefault("table-cell-properties/border-" + side) ?? p.GetValueOrDefault("table-cell-properties/border");
        if (value is null or "none") return Border.None;
        double width = 0.75;
        var kind = BorderStyle.Single;
        Color? color = null;
        foreach (var part in value.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (OdfUnits.Length(part) is { } w && char.IsDigit(part[0])) width = w;
            else if (part == "double") kind = BorderStyle.Double;
            else if (part == "dotted") kind = BorderStyle.Dotted;
            else if (part == "dashed") kind = BorderStyle.Dashed;
            else if (part == "hidden") return Border.None;
            else if (Color.TryParseHex(part) is { } c) color = c;
        }
        return new Border(kind, width, color ?? Color.Black);
    }

    private void ApplyPageSetup(Worksheet sheet, string? master)
    {
        if (styles.PageLayoutOf(master) is not { } layout || layout.Element(Ns.Style + "page-layout-properties") is not { } p) return;
        var setup = sheet.PageSetup;
        if (OdfUnits.Length((string?)p.Attribute(Ns.Fo + "page-width")) is { } w and > 0) setup.PageWidth = w;
        if (OdfUnits.Length((string?)p.Attribute(Ns.Fo + "page-height")) is { } h and > 0) setup.PageHeight = h;
        double Margin(string side, double fallback) => OdfUnits.Length((string?)p.Attribute(Ns.Fo + "margin-" + side)) ?? fallback;
        setup.Margins = new Thickness(Margin("top", 54), Margin("right", 54), Margin("bottom", 54), Margin("left", 54));
        if ((string?)p.Attribute(Ns.Style + "print-orientation") == "landscape") setup.Orientation = Orientation.Landscape;
        if ((string?)p.Attribute(Ns.Style + "print-page-order") == "ltr") setup.PageOrderOverThenDown = true;
        if ((string?)p.Attribute(Ns.Style + "table-centering") is "horizontal" or "both") setup.CenterHorizontally = true;
        if ((string?)p.Attribute(Ns.Style + "table-centering") is "vertical" or "both") setup.CenterVertically = true;
        if ((string?)p.Attribute(Ns.Style + "print") is { } print && print.Split(' ').Contains("grid")) setup.PrintGridLines = true;
        if (double.TryParse(((string?)p.Attribute(Ns.Style + "scale-to"))?.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var scale) && scale > 0 && scale <= 1000)
            setup.Scale = scale / 100;
        else if (int.TryParse((string?)p.Attribute(Ns.Style + "scale-to-pages"), out var pages) && pages > 0) { setup.FitToWidth = pages; setup.FitToHeight = pages; }
        if (int.TryParse((string?)p.Attribute(Ns.Style + "scale-to-X"), out var fx) && fx > 0) setup.FitToWidth = fx;
        if (int.TryParse((string?)p.Attribute(Ns.Style + "scale-to-Y"), out var fy) && fy > 0) setup.FitToHeight = fy;
    }

    private static int Repeat(XmlReader r, string name) =>
        int.TryParse(r.GetAttribute(name, Ns.Table.NamespaceName), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 1;

    private static int RepeatOf(XElement e, string name, XNamespace? ns = null) =>
        int.TryParse((string?)e.Attribute((ns ?? Ns.Table) + name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 1;
}
