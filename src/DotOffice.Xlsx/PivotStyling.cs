using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>
/// The look of a pivot table: Excel draws it from a built-in table style and from formats the user applied to areas of the table, and
/// neither is stored in the cells. This applies a few families of the built-in styles (approximately) and the pivot-area formats that
/// select whole parts of the table.
/// </summary>
internal static partial class PivotStyling
{
    /// <summary>What happened to the table's style and formats, for the warning.</summary>
    public sealed record Result(string? StyleName, StyleOutcome Style, int FormatsSkipped);

    public enum StyleOutcome { None, Approximated, NotApplied }

    [GeneratedRegex(@"^PivotStyle(Light|Medium|Dark)(\d{1,2})$")]
    private static partial Regex StyleName();

    /// <summary>The parts of the table a style or a format addresses.</summary>
    private sealed class Geometry
    {
        public CellRange Range;
        public int HeaderRows, LabelColumns;
        public Dictionary<int, string> RowTypes = [];     // sheet row -> "grand" | "default" (a subtotal)
        public Dictionary<int, string> ColumnTypes = [];

        public bool IsHeader(int row) => row - Range.FirstRow < HeaderRows;
        public bool IsLabel(int column) => column - Range.FirstColumn < LabelColumns;
        public bool IsGrandRow(int row) => RowTypes.GetValueOrDefault(row) == "grand";
        public bool IsGrandColumn(int column) => ColumnTypes.GetValueOrDefault(column) == "grand";
        public bool IsSubtotalRow(int row) => RowTypes.GetValueOrDefault(row) == "default";
    }

    public static Result Apply(Worksheet sheet, XElement definition, XlsxTheme theme, ConditionalFormatting.Environment env)
    {
        XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(e => e.Name.LocalName == name);
        bool Flag(XElement? e, string name, bool fallback) => (string?)e?.Attribute(name) switch { "1" or "true" => true, "0" or "false" => false, _ => fallback };
        int Int(XElement? e, string name, int fallback) => int.TryParse((string?)e?.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : fallback;

        var location = Child(definition, "location");
        var styleInfo = Child(definition, "pivotTableStyleInfo");
        string? styleName = (string?)styleInfo?.Attribute("name") is { Length: > 0 } n ? n : null;
        var formats = Child(definition, "formats")?.Elements().Where(e => e.Name.LocalName == "format").ToList() ?? [];
        if (location is null || !CellReference.TryParseRange(((string?)location.Attribute("ref") ?? "").AsSpan(), out var range) || (styleName is null && formats.Count == 0))
            return new Result(styleName, styleName is null ? StyleOutcome.None : StyleOutcome.NotApplied, formats.Count);

        var geometry = new Geometry
        {
            Range = range, HeaderRows = Math.Max(Int(location, "firstDataRow", 1), 0), LabelColumns = Math.Max(Int(location, "firstDataCol", 1), 0),
        };
        // Subtotal and grand-total rows and columns, from the table's item lists (or, without them, a grand total row at the bottom).
        void Types(string listName, Dictionary<int, string> into, int first, int last)
        {
            var items = Child(definition, listName)?.Elements().Where(e => e.Name.LocalName == "i").ToList();
            if (items is null) return;
            for (int k = 0; k < items.Count && first + k <= last; k++)
                if ((string?)items[k].Attribute("t") is { } t && t is "grand" or "default") into[first + k] = t;
        }
        Types("rowItems", geometry.RowTypes, range.FirstRow + geometry.HeaderRows, range.LastRow);
        Types("colItems", geometry.ColumnTypes, range.FirstColumn + geometry.LabelColumns, range.LastColumn);
        if (Child(definition, "rowItems") is null && Flag(definition, "rowGrandTotals", true) && range.LastRow - range.FirstRow >= geometry.HeaderRows)
            geometry.RowTypes[range.LastRow] = "grand";

        var outcome = StyleOutcome.None;
        if (styleName is not null)
        {
            var look = LookOf(styleName, theme);
            if (look is null) outcome = StyleOutcome.NotApplied;
            else
            {
                outcome = StyleOutcome.Approximated;
                ApplyLook(sheet, geometry, look, env, Flag(styleInfo, "showRowHeaders", false), Flag(styleInfo, "showColHeaders", false),
                    Flag(styleInfo, "showRowStripes", false), Flag(styleInfo, "showColStripes", false));
            }
        }

        int skipped = 0;
        foreach (var format in formats)
        {
            var area = Child(format, "pivotArea");
            int id = Int(format, "dxfId", -1);
            var test = area is null ? null : AreaTest(area, geometry);
            if (test is null || id < 0 || id >= env.Dxfs.Count) { skipped++; continue; }
            For(geometry, (row, column) => { if (test(row, column)) ConditionalFormatting.ApplyDxf(sheet, row, column, env.Dxfs[id], env); });
        }
        return new Result(styleName, outcome, skipped);
    }

    private static void For(Geometry g, Action<int, int> action)
    {
        for (int r = g.Range.FirstRow; r <= g.Range.LastRow; r++)
            for (int c = g.Range.FirstColumn; c <= g.Range.LastColumn; c++) action(r, c);
    }

    /// <summary>
    /// Which cells a pivot area selects, when it names whole parts of the table (everything, the corner, the column headings, the data, the
    /// labels, the grand totals, the field buttons). Areas that pick particular items through <c>references</c> give null.
    /// </summary>
    private static Func<int, int, bool>? AreaTest(XElement area, Geometry g)
    {
        if (area.Elements().Any(e => e.Name.LocalName == "references" && e.HasElements)) return null;
        string type = (string?)area.Attribute("type") ?? "normal";
        bool dataOnly = (string?)area.Attribute("dataOnly") is not ("0" or "false"), labelOnly = (string?)area.Attribute("labelOnly") is "1" or "true";
        bool grandRow = (string?)area.Attribute("grandRow") is "1" or "true", grandColumn = (string?)area.Attribute("grandCol") is "1" or "true";
        string? axis = (string?)area.Attribute("axis");
        switch (type)
        {
            case "all": return (_, _) => true;
            case "origin": return (r, c) => g.IsHeader(r) && g.IsLabel(c);
            case "topRight": return (r, c) => g.IsHeader(r) && !g.IsLabel(c);
            case "button":
                return axis == "axisCol" ? (r, c) => r == g.Range.FirstRow && c == g.Range.FirstColumn + g.LabelColumns
                    : (r, c) => r == g.Range.FirstRow + Math.Max(g.HeaderRows - 1, 0) && g.IsLabel(c);
            case "normal" or "data":
                if (grandRow || grandColumn)
                    return (r, c) => !g.IsHeader(r) && (grandRow && g.IsGrandRow(r) || grandColumn && g.IsGrandColumn(c)) && (labelOnly ? g.IsLabel(c) : !dataOnly || !g.IsLabel(c));
                if (axis is not null || area.Attribute("field") is not null) return null; // a field's labels: which items is a matter of the field
                return labelOnly ? (r, c) => !g.IsHeader(r) && g.IsLabel(c) : (r, c) => !g.IsHeader(r) && !g.IsLabel(c);
            default: return null;
        }
    }

    // ------------------------------------------------------------------ built-in styles

    /// <summary>The formats of the parts of a table that a style defines; null for a part it leaves alone.</summary>
    private sealed record Look(Dxf? Header, Dxf? RowHeader, Dxf? Subtotal, Dxf? GrandRow, Dxf? GrandColumn, Dxf? Stripe, Dxf? Body = null, Dxf? StripeAlternate = null);

    private static Border Line(Color color, bool thick = false) => new(thick ? BorderStyle.Thick : BorderStyle.Single, thick ? 2.25 : 0.75, color);

    /// <summary>
    /// The look of a built-in pivot style by its name, from its color (the accent of the theme the style's number selects, or black)
    /// and the family it belongs to. Only the families named below are drawn; the exact fills and borders of Excel's own definitions are
    /// not reproduced, so what is drawn follows their structure (what is highlighted) rather than every shade.
    /// </summary>
    private static Look? LookOf(string name, XlsxTheme theme)
    {
        var match = StyleName().Match(name);
        if (!match.Success) return null;
        string family = match.Groups[1].Value;
        int number = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        if (number is < 1 or > 28) return null;
        int accentIndex = (number - 1) % 7, variant = (number - 1) / 7;
        var accent = accentIndex == 0 ? Color.Black : theme[3 + accentIndex] ?? Color.Black;
        var tint = accentIndex == 0 ? new Color(0xD9, 0xD9, 0xD9) : accent.Blend(Color.White, 0.8);
        var light = accentIndex == 0 ? new Color(0xF2, 0xF2, 0xF2) : accent.Blend(Color.White, 0.9);
        var medium = accent.Blend(Color.White, 0.6);

        Dxf Bold(Color? fill = null, Color? font = null, Border? top = null, Border? bottom = null) =>
            new() { Bold = true, Fill = fill, FontColor = font, Top = top, Bottom = bottom };

        switch (family, variant)
        {
            case ("Light", 0): // Light 1-7: bold headings between thin lines
                return new Look(Bold(bottom: Line(accent)), null, Bold(), Bold(top: Line(accent), bottom: Line(accent)), Bold(), new Dxf { Fill = tint });
            case ("Light", 2): // Light 15-21 (Light 16 is Excel's default): a tinted header band, bold totals
                return new Look(Bold(fill: tint, top: Line(accent), bottom: Line(accent)), null, Bold(top: Line(accent)), Bold(top: Line(accent), bottom: Line(accent)), Bold(), new Dxf { Fill = light });
            case ("Medium", 0): // Medium 1-7 (Medium 2 is blue): a header in the style's color, banded rows, a bold grand total under a heavy line
                return new Look(Bold(fill: accent, font: Color.White), null, Bold(fill: tint), Bold(top: Line(accent, thick: true)), Bold(), new Dxf { Fill = tint });
            case ("Medium", 1): // Medium 8-14 (Medium 9 is blue): tinted body, a header and grand total in the style's color
                return new Look(Bold(fill: accent, font: Color.White), Bold(), Bold(fill: medium), Bold(fill: accent, font: Color.White), Bold(), new Dxf { Fill = light }, Body: new Dxf { Fill = tint }, StripeAlternate: new Dxf { Fill = light });
            default: return null;
        }
    }

    private static void ApplyLook(Worksheet sheet, Geometry g, Look look, ConditionalFormatting.Environment env, bool rowHeaders, bool columnHeaders, bool rowStripes, bool columnStripes)
    {
        void Put(int row, int column, Dxf? dxf)
        {
            if (dxf is not null) ConditionalFormatting.ApplyDxf(sheet, row, column, dxf, env);
        }
        For(g, (row, column) =>
        {
            bool header = g.IsHeader(row), label = g.IsLabel(column);
            int dataRow = row - g.Range.FirstRow - g.HeaderRows, dataColumn = column - g.Range.FirstColumn - g.LabelColumns;
            // From the bottom layer up: the body and its stripes, the row labels, subtotals and grand totals, then the header.
            if (!header && !label) Put(row, column, look.Body);
            // Bands run across whole rows (labels included) but skip the totals.
            bool banded = !header && rowStripes && !g.IsGrandRow(row) && !g.IsSubtotalRow(row);
            if (banded && dataRow % 2 == (look.StripeAlternate is null ? 0 : 1)) Put(row, column, look.Stripe);
            else if (banded && look.StripeAlternate is not null && !label) Put(row, column, look.StripeAlternate);
            if (!header && !label && columnStripes && dataColumn % 2 == 0 && look.Body is null) Put(row, column, look.Stripe);
            if (!header && label && rowHeaders) Put(row, column, look.RowHeader);
            if (!header && g.IsSubtotalRow(row)) Put(row, column, look.Subtotal);
            if (!header && g.IsGrandColumn(column)) Put(row, column, look.GrandColumn);
            if (!header && g.IsGrandRow(row)) Put(row, column, look.GrandRow);
            if (header && columnHeaders) Put(row, column, look.Header);
        });
    }
}
