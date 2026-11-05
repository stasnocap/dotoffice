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

    public static Result Apply(Worksheet sheet, XElement definition, XlsxTheme theme, ConditionalFormatting.Environment env,
        IReadOnlyDictionary<string, IReadOnlyList<(string Type, int DxfId)>>? customStyles = null)
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
            var look = customStyles is not null && customStyles.TryGetValue(styleName, out var elements) ? CustomLook(elements, env) : LookOf(styleName, theme);
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
    private sealed record Look(Dxf? Header, Dxf? RowHeader, Dxf? Subtotal, Dxf? GrandRow, Dxf? GrandColumn, Dxf? Stripe, Dxf? Body = null, Dxf? StripeAlternate = null)
    {
        /// <summary>Under everything else, on every cell (a custom style's <c>wholeTable</c>).</summary>
        public Dxf? Whole { get; init; }

        /// <summary>Column stripes of a custom style (the first, third… and the second, fourth… data column).</summary>
        public Dxf? ColumnStripe { get; init; }

        public Dxf? ColumnStripeAlternate { get; init; }

        /// <summary>A custom style: the stripes start with <see cref="Stripe"/> on the first data row and alternate with <see cref="StripeAlternate"/>.</summary>
        public bool Custom { get; init; }
    }

    /// <summary>
    /// The look a workbook's own table style gives: its elements name the dxf for the whole table, the header row, the total row and column,
    /// the first column, the subtotal rows and the row and column stripes.
    /// </summary>
    private static Look? CustomLook(IReadOnlyList<(string Type, int DxfId)> elements, ConditionalFormatting.Environment env)
    {
        Dxf? Of(string type)
        {
            int id = elements.Where(e => e.Type == type).Select(e => e.DxfId).DefaultIfEmpty(-1).First();
            return id >= 0 && id < env.Dxfs.Count ? env.Dxfs[id] : null;
        }
        var look = new Look(Of("headerRow"), Of("firstColumn"), Of("firstSubtotalRow"), Of("totalRow"), Of("lastColumn"), Of("firstRowStripe"), null, Of("secondRowStripe"))
        {
            Whole = Of("wholeTable"), ColumnStripe = Of("firstColumnStripe"), ColumnStripeAlternate = Of("secondColumnStripe"), Custom = true,
        };
        return look.Header is null && look.RowHeader is null && look.Subtotal is null && look.GrandRow is null && look.GrandColumn is null && look.Stripe is null
            && look.StripeAlternate is null && look.Whole is null && look.ColumnStripe is null && look.ColumnStripeAlternate is null ? null : look;
    }

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
        var dark = accent.Blend(Color.Black, 0.5);

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
            case ("Light", 1): // Light 8-14: a solid header in the style's color over banded rows, bold totals set off by lines
                return new Look(Bold(fill: accent, font: Color.White), null, Bold(top: Line(accent)), Bold(top: Line(accent), bottom: Line(accent)), Bold(), new Dxf { Fill = light });
            case ("Light", 3): // Light 22-28: like 15-21 with a stronger band in the body
                return new Look(Bold(fill: tint, bottom: Line(accent)), Bold(), Bold(top: Line(accent)), Bold(top: Line(accent, thick: true), bottom: Line(accent)), Bold(), new Dxf { Fill = tint });
            case ("Medium", 2): // Medium 15-21: a header in the style's color, a tinted body with darker bands, a bold grand total
                return new Look(Bold(fill: accent, font: Color.White), Bold(), Bold(fill: medium), Bold(top: Line(accent, thick: true)), Bold(), new Dxf { Fill = tint }) { Whole = new Dxf { Fill = light } };
            case ("Medium", 3): // Medium 22-28: header and grand total in the style's color, tinted row headings
                return new Look(Bold(fill: accent, font: Color.White), Bold(fill: tint), Bold(fill: medium), Bold(fill: accent, font: Color.White), Bold(), new Dxf { Fill = tint }) { Whole = new Dxf { Fill = light } };
            case ("Dark", _): // Dark 1-28: a header and grand total in a dark shade of the style's color over a tinted body with darker bands
                return new Look(Bold(fill: dark, font: Color.White), variant >= 2 ? Bold() : null, Bold(fill: medium), Bold(fill: dark, font: Color.White), Bold(), new Dxf { Fill = medium }) { Whole = new Dxf { Fill = tint } };
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
            // From the bottom layer up: the whole table, the body and its stripes, the row labels, subtotals and grand totals, then the header.
            Put(row, column, look.Whole);
            if (!header && !label) Put(row, column, look.Body);
            // Bands run across whole rows (labels included) but skip the totals.
            bool banded = !header && rowStripes && !g.IsGrandRow(row) && !g.IsSubtotalRow(row);
            if (banded && look.Custom) Put(row, column, dataRow % 2 == 0 ? look.Stripe : look.StripeAlternate);
            else if (banded && dataRow % 2 == (look.StripeAlternate is null ? 0 : 1)) Put(row, column, look.Stripe);
            else if (banded && look.StripeAlternate is not null && !label) Put(row, column, look.StripeAlternate);
            if (!header && !label && columnStripes && look.Custom) Put(row, column, dataColumn % 2 == 0 ? look.ColumnStripe : look.ColumnStripeAlternate);
            else if (!header && !label && columnStripes && dataColumn % 2 == 0 && look.Body is null) Put(row, column, look.Stripe);
            if (!header && label && rowHeaders) Put(row, column, look.RowHeader);
            if (!header && g.IsSubtotalRow(row)) Put(row, column, look.Subtotal);
            if (!header && g.IsGrandColumn(column)) Put(row, column, look.GrandColumn);
            if (!header && g.IsGrandRow(row)) Put(row, column, look.GrandRow);
            if (header && columnHeaders) Put(row, column, look.Header);
        });
    }
}
