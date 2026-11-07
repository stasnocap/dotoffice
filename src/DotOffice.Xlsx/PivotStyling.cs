using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Dml;

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
        public List<int> RowFields = [], ColumnFields = [];             // pivot field indexes by level (-2 is the data field)
        public Dictionary<int, int[]> RowPaths = [], ColumnPaths = [];  // sheet row/column -> the item index at each level of the axis
        public int DataFieldCount;
        public bool Compact = true;
        /// <summary>The page fields (filters) above the table, in order: the pivot field index and the item shown (-1 for "(All)").</summary>
        public List<(int Field, int Item)> PageFields = [];
        public int RowPageCount, ColumnPageCount;
        public bool PageOverThenDown;

        /// <summary>
        /// The cells of a page field: its name and the item shown. They lie above the table, separated from it by one empty row, in rows of
        /// <see cref="RowPageCount"/> fields filled down then over (or over then down) with a gap column between the columns of fields.
        /// </summary>
        public (int Row, int LabelColumn)? PageCell(int index)
        {
            int rows = Math.Max(RowPageCount, 1), columns = Math.Max(ColumnPageCount, 1);
            int down = PageOverThenDown ? index / columns : index % rows, over = PageOverThenDown ? index % columns : index / rows;
            int rowCount = PageOverThenDown ? (PageFields.Count + columns - 1) / columns : rows;
            int row = Range.FirstRow - 1 - rowCount + down;
            return row < 0 ? null : (row, Range.FirstColumn + over * 3);
        }

        public bool IsHeader(int row) => row - Range.FirstRow < HeaderRows;
        public bool IsLabel(int column) => column - Range.FirstColumn < LabelColumns;
        public bool IsGrandRow(int row) => RowTypes.GetValueOrDefault(row) == "grand";
        public bool IsGrandColumn(int column) => ColumnTypes.GetValueOrDefault(column) == "grand";
        public bool IsSubtotalRow(int row) => RowTypes.GetValueOrDefault(row) == "default";
    }

    public static Result Apply(Worksheet sheet, XElement definition, DmlTheme theme, ConditionalFormatting.Environment env,
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

        geometry.DataFieldCount = Child(definition, "dataFields")?.Elements().Count(e => e.Name.LocalName == "dataField") ?? 0;
        geometry.Compact = Flag(definition, "compact", true);
        geometry.RowPageCount = Math.Max(Int(location, "rowPageCount", 0), 0);
        geometry.ColumnPageCount = Math.Max(Int(location, "colPageCount", 0), 0);
        geometry.PageOverThenDown = Flag(definition, "pageOverThenDown", false);
        geometry.PageFields = Child(definition, "pageFields")?.Elements().Where(e => e.Name.LocalName == "pageField").Select(e => (Int(e, "fld", -1), Int(e, "item", -1))).ToList() ?? [];
        if (geometry.RowPageCount == 0 && geometry.PageFields.Count > 0) geometry.RowPageCount = geometry.PageFields.Count;
        geometry.RowFields = Fields("rowFields");
        geometry.ColumnFields = Fields("colFields");
        Paths("rowItems", geometry.RowPaths, range.FirstRow + geometry.HeaderRows, range.LastRow);
        Paths("colItems", geometry.ColumnPaths, range.FirstColumn + geometry.LabelColumns, range.LastColumn);
        List<int> Fields(string name) => Child(definition, name)?.Elements().Where(e => e.Name.LocalName == "field").Select(e => Int(e, "x", -1)).ToList() ?? [];
        // The item list gives each row (column) the items it sits under: `r` leading entries repeat the previous row's, the rest are listed.
        void Paths(string listName, Dictionary<int, int[]> into, int first, int last)
        {
            var items = Child(definition, listName)?.Elements().Where(e => e.Name.LocalName == "i").ToList();
            if (items is null) return;
            int[] previous = [];
            for (int k = 0; k < items.Count && first + k <= last; k++)
            {
                if ((string?)items[k].Attribute("t") is "grand" or "data" or "blank") { previous = []; continue; }
                var listed = items[k].Elements().Where(e => e.Name.LocalName == "x").Select(e => Int(e, "v", 0));
                var path = previous.Take(Math.Min(Int(items[k], "r", 0), previous.Length)).Concat(listed).ToArray();
                into[first + k] = path;
                previous = path;
            }
        }

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
            if (test is null && area is not null && PageTest(area, geometry) is not null) test = (_, _) => false;
            if (test is null || id < 0 || id >= env.Dxfs.Count) { skipped++; continue; }
            For(geometry, (row, column) => { if (test(row, column)) ConditionalFormatting.ApplyDxf(sheet, row, column, env.Dxfs[id], env); });
            // The page fields lie above the table, outside its range.
            if (PageTest(area!, geometry) is { } pageTest)
                for (int k = 0; k < geometry.PageFields.Count; k++)
                    if (geometry.PageCell(k) is var (pageRow, pageColumn))
                        foreach (int column in new[] { pageColumn, pageColumn + 1 })
                            if (pageTest(k, column == pageColumn)) ConditionalFormatting.ApplyDxf(sheet, pageRow, column, env.Dxfs[id], env);
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
    /// <summary>
    /// A pivot area on the page fields: the field's button (its name) with <c>axis="axisPage"</c>, or a <c>reference</c> to a page field and
    /// the item shown (the cell with the item). Gives which of the cells of the field with a given index it covers (label cell or item cell).
    /// </summary>
    private static Func<int, bool, bool>? PageTest(XElement area, Geometry g)
    {
        if (g.PageFields.Count == 0) return null;
        string type = (string?)area.Attribute("type") ?? "normal";
        var references = area.Elements().FirstOrDefault(e => e.Name.LocalName == "references")?.Elements().Where(e => e.Name.LocalName == "reference").ToList() ?? [];
        long.TryParse((string?)area.Attribute("field"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long areaField);
        if ((string?)area.Attribute("axis") == "axisPage" && references.Count == 0)
        {
            if (type != "button") return area.Attribute("field") is null ? null : (index, label) => g.PageFields[index].Field == areaField && !label;
            return (index, label) => label && (area.Attribute("field") is null || g.PageFields[index].Field == areaField);
        }
        if (references.Count == 0) return null;
        // References to page fields only: the item cell of each, when the item named is the one shown (or the field shows all).
        var picks = new List<(int Field, HashSet<int> Items)>();
        foreach (var reference in references)
        {
            if (!int.TryParse((string?)reference.Attribute("field"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int field) || g.PageFields.All(p => p.Field != field)) return null;
            picks.Add((field, reference.Elements().Where(e => e.Name.LocalName == "x").Select(e => int.TryParse((string?)e.Attribute("v"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : 0).ToHashSet()));
        }
        bool labelOnly = (string?)area.Attribute("labelOnly") is "1" or "true";
        return (index, label) =>
        {
            var (field, item) = g.PageFields[index];
            return !label && !labelOnly && picks.Any(p => p.Field == field && (item < 0 || p.Items.Contains(item)));
        };
    }

    private static Func<int, int, bool>? AreaTest(XElement area, Geometry g)
    {
        if (area.Elements().FirstOrDefault(e => e.Name.LocalName == "references" && e.HasElements) is { } references) return ItemTest(area, references, g);
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
                if (axis == "axisPage") return null;
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

    /// <summary>
    /// A pivot area that picks items: each <c>reference</c> names a pivot field and the items of it (indexes into the field's item list,
    /// the same numbers the table's item lists use). The rows or columns whose path holds those items are selected; their data cells, their
    /// labels or both, by <c>dataOnly</c> and <c>labelOnly</c>. Only fields on the row or column axis are understood (a data-field reference
    /// counts when there is a single data field); anything else gives null, which the caller reports.
    /// </summary>
    private static Func<int, int, bool>? ItemTest(XElement area, XElement references, Geometry g)
    {
        string type = (string?)area.Attribute("type") ?? "normal";
        if (type is not ("normal" or "data")) return null;
        bool labelOnly = (string?)area.Attribute("labelOnly") is "1" or "true", dataOnly = (string?)area.Attribute("dataOnly") is not ("0" or "false");
        var rowPicks = new List<(int Level, HashSet<int> Items)>();
        var columnPicks = new List<(int Level, HashSet<int> Items)>();
        foreach (var reference in references.Elements().Where(e => e.Name.LocalName == "reference"))
        {
            // Excel writes the data field as the unsigned 4294967294.
            if (!long.TryParse((string?)reference.Attribute("field"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long wide)) return null;
            int field = wide == 4294967294 ? -2 : (int)wide;
            // With several data fields the data-field "field" is a level of an axis (its items are the data fields' indexes); with one it adds no condition.
            if (field == -2 && g.DataFieldCount <= 1) continue;
            var items = reference.Elements().Where(e => e.Name.LocalName == "x").Select(e => int.TryParse((string?)e.Attribute("v"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : 0).ToHashSet();
            if (items.Count == 0) return null;
            int level = g.RowFields.IndexOf(field);
            if (level >= 0) { rowPicks.Add((level, items)); continue; }
            level = g.ColumnFields.IndexOf(field);
            if (level < 0) return null;
            columnPicks.Add((level, items));
        }
        if (rowPicks.Count == 0 && columnPicks.Count == 0) return null;
        bool Matches(int[]? path, List<(int Level, HashSet<int> Items)> picks) =>
            picks.Count == 0 || path is not null && picks.All(p => p.Level < path.Length && p.Items.Contains(path[p.Level]));
        // The row's own level of the deepest pick decides where its label sits: column 0 when the table is compact, else the field's column.
        int rowLevel = rowPicks.Count == 0 ? -1 : rowPicks.Max(p => p.Level), columnLevel = columnPicks.Count == 0 ? -1 : columnPicks.Max(p => p.Level);
        bool Heading(int[]? path, int level, List<(int Level, HashSet<int> Items)> picks, Dictionary<int, int[]> all, int position)
        {
            if (level < 0) return true;
            if (path is null) return false;
            // The label of the item shows on the first row of its group (the row where the item at that level starts).
            return path.Length == level + 1 || !all.TryGetValue(position - 1, out var before) || before.Length <= level || before[level] != path[level];
        }
        return (r, c) =>
        {
            bool header = g.IsHeader(r), label = g.IsLabel(c);
            g.RowPaths.TryGetValue(r, out var rowPath);
            g.ColumnPaths.TryGetValue(c, out var columnPath);
            if (rowPicks.Count > 0 && (header || !Matches(rowPath, rowPicks))) return false;
            if (columnPicks.Count > 0 && (label || !Matches(columnPath, columnPicks))) return false;
            if (header && columnPicks.Count == 0 || label && rowPicks.Count == 0) return false;
            bool isLabel = header || label;
            if (isLabel)
            {
                if (dataOnly && !labelOnly) return false; // the default selects data cells only
                if (label) // a row label: the field's column, on the group's first row
                {
                    if (!Heading(rowPath, rowLevel, rowPicks, g.RowPaths, r)) return false;
                    return g.Compact ? c == g.Range.FirstColumn : c - g.Range.FirstColumn == Math.Min(rowLevel, g.LabelColumns - 1);
                }
                return Heading(columnPath, columnLevel, columnPicks, g.ColumnPaths, c) && r - g.Range.FirstRow == Math.Min(columnLevel + (g.HeaderRows > 1 ? 1 : 0), g.HeaderRows - 1);
            }
            return !labelOnly;
        };
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
    private static Look? LookOf(string name, DmlTheme theme)
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
