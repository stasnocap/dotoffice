using System.Globalization;
using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxConditionalFormattingTests
{
    private const string Red = "FFC7CE", Green = "C6EFCE", Blue = "0000FF";

    /// <summary>Numbers down column A (or the given column), starting at row 1; null entries leave a gap.</summary>
    private static string Column(string column, params double?[] values)
    {
        var rows = new System.Text.StringBuilder();
        for (int i = 0; i < values.Length; i++)
            if (values[i] is { } v) rows.Append(Row(i + 1, C($"{column}{i + 1}", v.ToString(CultureInfo.InvariantCulture))));
        return rows.ToString();
    }

    private static Worksheet Read(string rows, string tail, string[]? dxfs = null, ConversionWarningList? warnings = null, string? definedNames = null, DateTime? today = null)
    {
        var builder = new XlsxBuilder { Styles = StylesWithDxfs(dxfs ?? [Dxf(Red), Dxf(Green)]), DefinedNames = definedNames }.AddSheet("S", rows, tail: tail);
        return XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings, Today = today }).Sheets[0];
    }

    private static Color Hex(string hex) => Color.TryParseHex(hex)!.Value;

    private static string FilledRows(Worksheet sheet, string? hex = null) =>
        string.Join(",", sheet.Cells.Where(kv => kv.Value.Style.Fill == (hex is null ? kv.Value.Style.Fill : Hex(hex)) && kv.Value.Style.Fill is not null)
            .OrderBy(kv => kv.Key.Row).ThenBy(kv => kv.Key.Column).Select(kv => $"{(char)('A' + kv.Key.Column)}{kv.Key.Row + 1}"));

    [Fact]
    public void Cells_keep_their_raw_value()
    {
        var sheet = Read(Column("A", 12.5), "");
        Assert.Equal(12.5, sheet.Cells[(0, 0)].Value);
    }

    [Fact]
    public void CellIs_greater_than_fills_matching_cells_only()
    {
        var sheet = Read(Column("A", 1, 10, 5), ConditionalFormatting("A1:A3", CellIs("greaterThan", "4", 1, 0)));
        Assert.Equal("A2,A3", FilledRows(sheet, Red));
        Assert.Equal("10", sheet.Cells[(1, 0)].Text);
        Assert.Null(sheet.Cells[(0, 0)].Style.Fill);
    }

    [Theory]
    [InlineData("lessThan", "5", "A1")]
    [InlineData("lessThanOrEqual", "5", "A1,A2")]
    [InlineData("equal", "5", "A2")]
    [InlineData("notEqual", "5", "A1,A3")]
    [InlineData("greaterThanOrEqual", "5", "A2,A3")]
    [InlineData("greaterThan", "5", "A3")]
    public void CellIs_operators(string op, string operand, string expected)
    {
        var sheet = Read(Column("A", 1, 5, 9), ConditionalFormatting("A1:A3", CellIs(op, operand, 1, 0)));
        Assert.Equal(expected, FilledRows(sheet));
    }

    [Fact]
    public void CellIs_between_and_not_between_and_a_cell_reference_operand()
    {
        var rows = Column("A", 1, 5, 9) + Row(1, C("A1", "1"), C("C1", "6"));
        var between = Read(rows, ConditionalFormatting("A1:A3", CellIs("between", "2", 1, 0, second: "8")));
        Assert.Equal("A2", FilledRows(between));
        var notBetween = Read(rows, ConditionalFormatting("A1:A3", CellIs("notBetween", "2", 1, 0, second: "8")));
        Assert.Equal("A1,A3", FilledRows(notBetween));
        // The operand $C$1 is 6: values above it.
        var reference = Read(rows, ConditionalFormatting("A1:A3", CellIs("greaterThan", "$C$1", 1, 0)));
        Assert.Equal("A3", FilledRows(reference));
    }

    [Fact]
    public void CellIs_compares_text_without_regard_to_case()
    {
        var sheet = Read(Row(1, Inline("A1", "Done")) + Row(2, Inline("A2", "open")), ConditionalFormatting("A1:A2", CellIs("equal", "\"DONE\"", 1, 0)));
        Assert.Equal("A1", FilledRows(sheet));
    }

    [Fact]
    public void Higher_priority_wins_conflicting_properties_and_the_rest_merges()
    {
        var dxfs = new[] { Dxf(Red), Dxf(Green, font: "<b/>") };
        var sheet = Read(Column("A", 10), ConditionalFormatting("A1", CellIs("greaterThan", "5", 2, 1), CellIs("greaterThan", "0", 1, 0)), dxfs);
        var cell = sheet.Cells[(0, 0)];
        Assert.Equal(Hex(Red), cell.Style.Fill);
        Assert.True(cell.Style.Font.Bold);
    }

    [Fact]
    public void Priority_is_global_across_blocks_not_document_order()
    {
        var sheet = Read(Column("A", 10),
            ConditionalFormatting("A1", CellIs("greaterThan", "0", 2, 1)) + ConditionalFormatting("A1", CellIs("greaterThan", "0", 1, 0)));
        Assert.Equal(Hex(Red), sheet.Cells[(0, 0)].Style.Fill);
    }

    [Fact]
    public void StopIfTrue_hides_lower_priority_rules()
    {
        var dxfs = new[] { Dxf(Red), Dxf(null, font: "<b/>") };
        var stopped = Read(Column("A", 10), ConditionalFormatting("A1", CellIs("greaterThan", "5", 1, 0, stopIfTrue: true), CellIs("greaterThan", "0", 2, 1)), dxfs);
        Assert.False(stopped.Cells[(0, 0)].Style.Font.Bold);
        var notStopped = Read(Column("A", 10), ConditionalFormatting("A1", CellIs("greaterThan", "5", 1, 0), CellIs("greaterThan", "0", 2, 1)), dxfs);
        Assert.True(notStopped.Cells[(0, 0)].Style.Font.Bold);
        // A rule that does not match does not stop anything.
        var unmatched = Read(Column("A", 3), ConditionalFormatting("A1", CellIs("greaterThan", "5", 1, 0, stopIfTrue: true), CellIs("greaterThan", "0", 2, 1)), dxfs);
        Assert.True(unmatched.Cells[(0, 0)].Style.Font.Bold);
    }

    [Fact]
    public void Expression_with_relative_references_follows_the_top_left_of_the_range()
    {
        // Written for B2: every cell compares itself with the cell to its left.
        var rows = Row(2, C("A2", "5"), C("B2", "9")) + Row(3, C("A3", "5"), C("B3", "1")) + Row(4, C("A4", "5"), C("B4", "7"));
        var sheet = Read(rows, ConditionalFormatting("B2:B4", Expression("B2>A2", 1, 0)));
        Assert.Equal("B2,B4", FilledRows(sheet));
    }

    [Fact]
    public void Expression_with_mixed_absolute_references()
    {
        var rows = Row(1, C("A1", "1"), C("B1", "1")) + Row(2, C("A2", "1"), C("B2", "10")) + Row(3, C("A3", "1"), C("B3", "3"));
        var sheet = Read(rows, ConditionalFormatting("A1:A3", Expression("$B1>5", 1, 0)));
        Assert.Equal("A2", FilledRows(sheet));
    }

    [Fact]
    public void Expression_functions_and_operators()
    {
        string Filled(string formula, params double?[] values) =>
            FilledRows(Read(Column("A", values), ConditionalFormatting("A1:A4", Expression(formula, 1, 0))));
        Assert.Equal("A2,A3", Filled("AND(A1>1,A1<5)", 1, 2, 4, 9));
        Assert.Equal("A1,A4", Filled("OR(A1<2,A1>8)", 1, 2, 4, 9));
        Assert.Equal("A1,A2", Filled("NOT(A1>3)", 1, 2, 4, 9));
        Assert.Equal("A2,A4", Filled("MOD(A1,2)=0", 1, 2, 3, 4));
        Assert.Equal("A2,A4", Filled("MOD(ROW(),2)=0", 1, 1, 1, 1));
        Assert.Equal("A1,A3", Filled("MOD(COLUMN()+ROW(),2)=0", 1, 1, 1, 1));
        Assert.Equal("A2", Filled("A1=2*A1-2+(A1^2-A1*A1)", 1, 2, 3, 4));
        Assert.Equal("A3", Filled("ISBLANK(A1)", 1, 2, null, 4));
        Assert.Equal("A1,A2,A4", Filled("NOT(ISBLANK(A1))", 1, 2, null, 4));
        Assert.Equal("A1,A3", Filled("COUNTIF($A$1:$A$4,A1)>1", 7, 1, 7, 2));
        Assert.Equal("A4", Filled("A1>AVERAGE($A$1:$A$4)", 1, 2, 3, 100));
        Assert.Equal("A2", Filled("IF(A1>2,A1<5,FALSE)", 1, 3, 9, 1));
        Assert.Equal("A1", Filled("-A1<0", 1, -2, null, -4).Split(',')[0]);
        Assert.Equal("A2", Filled("A1&\"x\"=\"2x\"", 1, 2, 3, 4));
    }

    [Fact]
    public void Expression_errors_do_not_match()
    {
        var sheet = Read(Column("A", 1, 0), ConditionalFormatting("A1:A2", Expression("1/A1>0", 1, 0)));
        Assert.Equal("A1", FilledRows(sheet));
    }

    [Fact]
    public void Empty_cells_inside_the_used_range_get_fills_for_zebra_stripes()
    {
        var sheet = Read(Column("A", 1, 1, 1, 1) + Row(5, C("A5", "1")), ConditionalFormatting("A1:A5", Expression("MOD(ROW(),2)=0", 1, 0)));
        Assert.Equal("A2,A4", FilledRows(sheet));
        var gap = Read(Row(1, C("A1", "1"), C("C1", "1")), ConditionalFormatting("A1:C1", Expression("ISBLANK(A1)", 1, 0)));
        Assert.Equal("B1", FilledRows(gap));
    }

    [Fact]
    public void Ranges_are_clamped_to_the_used_range_and_extended_by_the_print_area()
    {
        var rows = Column("A", 1, 1, 1, 1);
        var stripes = ConditionalFormatting("A:B", Expression("MOD(ROW(),2)=0", 1, 0));
        var clamped = Read(rows, stripes);
        Assert.Equal("A2,A4", FilledRows(clamped));
        var area = Read(rows, stripes, definedNames: "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">S!$A$1:$B$6</definedName>");
        Assert.Equal("A2,B2,A4,B4,A6,B6", FilledRows(area, Red).Replace("A2,A4,A6,B2,B4,B6", "A2,B2,A4,B4,A6,B6"));
    }

    [Fact]
    public void Top_and_bottom_n_and_percent()
    {
        string Filled(string attributes, params double?[] values) =>
            FilledRows(Read(Column("A", values), ConditionalFormatting("A1:A5", CfRule("top10", 1, attributes, "", 0))));
        Assert.Equal("A4,A5", Filled("rank=\"2\"", 1, 2, 3, 4, 5));
        Assert.Equal("A1,A2", Filled("rank=\"2\" bottom=\"1\"", 1, 2, 3, 4, 5));
        Assert.Equal("A4,A5", Filled("rank=\"40\" percent=\"1\"", 1, 2, 3, 4, 5));
        Assert.Equal("A1,A2", Filled("rank=\"1\"", 5, 5, 4, 3, 2));
        Assert.Equal("A5", Filled("rank=\"1\" bottom=\"1\"", 5, 5, 4, 3, 2).Replace("A5", "A5"));
    }

    [Fact]
    public void Above_and_below_average()
    {
        string Filled(string attributes) =>
            FilledRows(Read(Column("A", 1, 2, 3, 4, 10), ConditionalFormatting("A1:A5", CfRule("aboveAverage", 1, attributes, "", 0))));
        Assert.Equal("A5", Filled(""));
        Assert.Equal("A1,A2,A3", Filled("aboveAverage=\"0\""));
        Assert.Equal("A4,A5", Filled("equalAverage=\"1\""));
        Assert.Equal("A5", Filled("stdDev=\"1\""));
        Assert.Equal("", Filled("aboveAverage=\"0\" stdDev=\"2\""));
    }

    [Fact]
    public void Duplicate_and_unique_values()
    {
        var rows = Row(1, Inline("A1", "a")) + Row(2, Inline("A2", "A")) + Row(3, Inline("A3", "b")) + Row(4, C("A4", "1")) + Row(5, C("A5", "1")) + Row(6, C("A6", "2"));
        Assert.Equal("A1,A2,A4,A5", FilledRows(Read(rows, ConditionalFormatting("A1:A6", CfRule("duplicateValues", 1, "", "", 0)))));
        Assert.Equal("A3,A6", FilledRows(Read(rows, ConditionalFormatting("A1:A6", CfRule("uniqueValues", 1, "", "", 0)))));
    }

    [Fact]
    public void Text_blank_and_error_rules()
    {
        var rows = Row(1, Inline("A1", "Alpha")) + Row(2, Inline("A2", "beta")) + Row(3, C("A3", "#N/A", "e")) + Row(4, C("B4", "1"));
        string Filled(string type, string attributes = "") => FilledRows(Read(rows, ConditionalFormatting("A1:A4", CfRule(type, 1, attributes, "", 0))));
        Assert.Equal("A1", Filled("containsText", "operator=\"containsText\" text=\"ALP\""));
        Assert.Equal("A2", Filled("beginsWith", "operator=\"beginsWith\" text=\"be\""));
        Assert.Equal("A2", Filled("endsWith", "operator=\"endsWith\" text=\"ta\""));
        Assert.Equal("A4", Filled("containsBlanks"));
        Assert.Equal("A3", Filled("containsErrors"));
    }

    [Fact]
    public void Color_scales_interpolate_between_the_colors()
    {
        string Scale(string cfvos, string colors) => CfRule("colorScale", 1, "", $"<colorScale>{cfvos}{colors}</colorScale>");
        var two = Read(Column("A", 0, 50, 100), ConditionalFormatting("A1:A3", Scale(Cfvo("min") + Cfvo("max"), Rgb("FFFFFF") + Rgb("FF0000"))));
        Assert.Equal(Hex("FFFFFF"), two.Cells[(0, 0)].Style.Fill);
        Assert.Equal(Hex("FF0000"), two.Cells[(2, 0)].Style.Fill);
        var middle = two.Cells[(1, 0)].Style.Fill!.Value;
        Assert.Equal(255, middle.R);
        Assert.InRange(middle.G, 126, 129);

        var three = Read(Column("A", 1, 2, 3, 4, 5), ConditionalFormatting("A1:A5",
            Scale(Cfvo("min") + Cfvo("percentile", "50") + Cfvo("max"), Rgb("FF0000") + Rgb("FFFF00") + Rgb("00FF00"))));
        Assert.Equal(Hex("FF0000"), three.Cells[(0, 0)].Style.Fill);
        Assert.Equal(Hex("FFFF00"), three.Cells[(2, 0)].Style.Fill);
        Assert.Equal(Hex("00FF00"), three.Cells[(4, 0)].Style.Fill);
        Assert.Equal(255, three.Cells[(1, 0)].Style.Fill!.Value.R);
        Assert.InRange(three.Cells[(1, 0)].Style.Fill!.Value.G, 126, 129);
    }

    [Fact]
    public void Color_scale_thresholds_clamp_and_skip_non_numbers()
    {
        var rows = Column("A", 0, 10, 20, 30) + Row(5, Inline("A5", "text"));
        var sheet = Read(rows, ConditionalFormatting("A1:A5", CfRule("colorScale", 1, "", $"<colorScale>{Cfvo("num", "10")}{Cfvo("num", "20")}{Rgb("000000")}{Rgb("FFFFFF")}</colorScale>")));
        Assert.Equal(Hex("000000"), sheet.Cells[(0, 0)].Style.Fill);
        Assert.Equal(Hex("000000"), sheet.Cells[(1, 0)].Style.Fill);
        Assert.Equal(Hex("FFFFFF"), sheet.Cells[(2, 0)].Style.Fill);
        Assert.Equal(Hex("FFFFFF"), sheet.Cells[(3, 0)].Style.Fill);
        Assert.Null(sheet.Cells[(4, 0)].Style.Fill);
    }

    [Fact]
    public void A_dxf_fill_with_higher_priority_beats_a_color_scale()
    {
        var scale = CfRule("colorScale", 2, "", $"<colorScale>{Cfvo("min")}{Cfvo("max")}{Rgb("000000")}{Rgb("FFFFFF")}</colorScale>");
        var sheet = Read(Column("A", 0, 10), ConditionalFormatting("A1:A2", CellIs("greaterThan", "5", 1, 0), scale));
        Assert.Equal(Hex(Red), sheet.Cells[(1, 0)].Style.Fill);
        Assert.Equal(Hex("000000"), sheet.Cells[(0, 0)].Style.Fill);
    }

    [Fact]
    public void Data_bars_use_the_ten_to_ninety_percent_range_by_default()
    {
        var sheet = Read(Column("A", 0, 50, 100), ConditionalFormatting("A1:A3", DataBar(1, attributes: "minLength=\"10\" maxLength=\"90\"")));
        var bars = Enumerable.Range(0, 3).Select(i => sheet.Cells[(i, 0)].DataBar!).ToList();
        Assert.Equal(0.1, bars[0].End, 6);
        Assert.Equal(0.5, bars[1].End, 6);
        Assert.Equal(0.9, bars[2].End, 6);
        Assert.All(bars, b => { Assert.Equal(0, b.Start); Assert.Equal(Hex("638EC6"), b.Color); Assert.True(b.Gradient); });
        Assert.Equal("50", sheet.Cells[(1, 0)].Text);
    }

    [Fact]
    public void Data_bar_gradient_flag_comes_from_the_matching_x14_extension()
    {
        const string solid = "{11111111-1111-1111-1111-111111111111}", fading = "{22222222-2222-2222-2222-222222222222}";
        var tail = ConditionalFormatting("A1:A2", DataBar(1, extensionId: solid)) + ConditionalFormatting("B1:B2", DataBar(2, extensionId: fading)) +
                   DataBarExtension(solid, gradient: false) + DataBarExtension(fading, gradient: true);
        var sheet = Read(Column("A", 1, 2) + Row(1, C("B1", "1")) + Row(2, C("B2", "2")), tail);
        Assert.False(sheet.Cells[(1, 0)].DataBar!.Gradient);
        Assert.True(sheet.Cells[(1, 1)].DataBar!.Gradient);
        // The extension says minLength 0 / maxLength 100.
        Assert.Equal(0, sheet.Cells[(0, 0)].DataBar!.End, 6);
        Assert.Equal(1, sheet.Cells[(1, 0)].DataBar!.End, 6);
    }

    [Fact]
    public void Data_bar_borders_come_from_the_extension()
    {
        const string id = "{44444444-4444-4444-4444-444444444444}";
        var tail = ConditionalFormatting("A1:A3", DataBar(1, extensionId: id)) +
                   DataBarExtension(id, gradient: false, "border=\"1\" negativeBarBorderColorSameAsPositive=\"0\"",
                       "<x14:negativeFillColor rgb=\"FF00FF00\"/><x14:borderColor rgb=\"FF0000FF\"/><x14:negativeBorderColor rgb=\"FFFF00FF\"/>");
        var sheet = Read(Column("A", -50, 0, 100), tail);
        Assert.Equal(Hex("FF00FF"), sheet.Cells[(0, 0)].DataBar!.BorderColor);
        Assert.Equal(Hex("0000FF"), sheet.Cells[(2, 0)].DataBar!.BorderColor);
        // Without border="1" there is no border, and without a border color it takes the bar's color.
        var plain = Read(Column("A", 1, 2), ConditionalFormatting("A1:A2", DataBar(1, extensionId: id)) + DataBarExtension(id, gradient: false));
        Assert.Null(plain.Cells[(1, 0)].DataBar!.BorderColor);
        var defaulted = Read(Column("A", 1, 2), ConditionalFormatting("A1:A2", DataBar(1, extensionId: id)) + DataBarExtension(id, gradient: false, "border=\"1\""));
        Assert.Equal(defaulted.Cells[(1, 0)].DataBar!.Color, defaulted.Cells[(1, 0)].DataBar!.BorderColor);
    }

    [Fact]
    public void Data_bars_with_negative_values_grow_from_an_axis()
    {
        const string id = "{33333333-3333-3333-3333-333333333333}";
        var tail = ConditionalFormatting("A1:A3", DataBar(1, extensionId: id)) + DataBarExtension(id, gradient: false, children: "<x14:negativeFillColor rgb=\"FF00FF00\"/>");
        var sheet = Read(Column("A", -50, 0, 100), tail);
        var negative = sheet.Cells[(0, 0)].DataBar!;
        Assert.True(negative.Negative);
        Assert.Equal(0, negative.Start, 6);
        Assert.Equal(1.0 / 3, negative.End, 6);
        Assert.Equal(Hex("00FF00"), negative.Color);
        var positive = sheet.Cells[(2, 0)].DataBar!;
        Assert.Equal(1.0 / 3, positive.Start, 6);
        Assert.Equal(1, positive.End, 6);
        Assert.Equal(Hex("638EC6"), positive.Color);
    }

    [Fact]
    public void Data_bar_can_hide_the_value()
    {
        var sheet = Read(Column("A", 1, 2), ConditionalFormatting("A1:A2", CfRule("dataBar", 1, "", $"<dataBar showValue=\"0\">{Cfvo("min")}{Cfvo("max")}{Rgb("638EC6")}</dataBar>")));
        Assert.Equal("", sheet.Cells[(0, 0)].Text);
        Assert.NotNull(sheet.Cells[(0, 0)].DataBar);
    }

    private static string Thresholds(params int[] percents) => string.Concat(percents.Select(p => Cfvo("percent", p.ToString())));

    [Fact]
    public void Icon_sets_pick_the_icon_by_threshold()
    {
        var values = Column("A", 1, 4, 8, 10);
        var arrows = Read(values, ConditionalFormatting("A1:A4", IconSet(1, "3Arrows", Thresholds(0, 33, 67))));
        Assert.Equal([IconShape.ArrowDown, IconShape.ArrowSideways, IconShape.ArrowUp, IconShape.ArrowUp], Enumerable.Range(0, 4).Select(i => arrows.Cells[(i, 0)].Icon!.Shape));
        Assert.Equal("4", arrows.Cells[(1, 0)].Text);
        var lights = Read(values, ConditionalFormatting("A1:A4", IconSet(1, "3TrafficLights1", Thresholds(0, 33, 67))));
        Assert.All(Enumerable.Range(0, 4), i => Assert.Equal(IconShape.Circle, lights.Cells[(i, 0)].Icon!.Shape));
        Assert.NotEqual(lights.Cells[(0, 0)].Icon!.Color, lights.Cells[(2, 0)].Icon!.Color);
        var symbols = Read(values, ConditionalFormatting("A1:A4", IconSet(1, "3Symbols", Thresholds(0, 33, 67))));
        Assert.Equal([IconShape.Cross, IconShape.Exclamation, IconShape.Check], new[] { 0, 1, 2 }.Select(i => symbols.Cells[(i, 0)].Icon!.Shape));
    }

    [Theory]
    [InlineData("3ArrowsGray", 3, IconShape.ArrowDown, IconShape.ArrowUp)]
    [InlineData("3Flags", 3, IconShape.Flag, IconShape.Flag)]
    [InlineData("3TrafficLights2", 3, IconShape.RimmedCircle, IconShape.RimmedCircle)]
    [InlineData("3Signs", 3, IconShape.Diamond, IconShape.Circle)]
    [InlineData("3Symbols2", 3, IconShape.BareCross, IconShape.BareCheck)]
    [InlineData("3Stars", 3, IconShape.Star, IconShape.Star)]
    [InlineData("3Triangles", 3, IconShape.TriangleDown, IconShape.TriangleUp)]
    [InlineData("4Arrows", 4, IconShape.ArrowDown, IconShape.ArrowUp)]
    [InlineData("4ArrowsGray", 4, IconShape.ArrowDown, IconShape.ArrowUp)]
    [InlineData("4RedToBlack", 4, IconShape.Circle, IconShape.Circle)]
    [InlineData("4Rating", 4, IconShape.Bars, IconShape.Bars)]
    [InlineData("4TrafficLights", 4, IconShape.Circle, IconShape.Circle)]
    [InlineData("5Arrows", 5, IconShape.ArrowDown, IconShape.ArrowUp)]
    [InlineData("5ArrowsGray", 5, IconShape.ArrowDown, IconShape.ArrowUp)]
    [InlineData("5Rating", 5, IconShape.Bars, IconShape.Bars)]
    [InlineData("5Quarters", 5, IconShape.Quarters, IconShape.Quarters)]
    [InlineData("5Boxes", 5, IconShape.Boxes, IconShape.Boxes)]
    public void All_icon_sets_are_read(string name, int count, IconShape lowest, IconShape highest)
    {
        var thresholds = Thresholds(Enumerable.Range(0, count).Select(i => i * 100 / count).ToArray());
        var warnings = new ConversionWarningList();
        var sheet = Read(Column("A", 0, 100), ConditionalFormatting("A1:A2", IconSet(1, name, thresholds)), warnings: warnings);
        Assert.Empty(warnings.Warnings);
        Assert.Equal(lowest, sheet.Cells[(0, 0)].Icon!.Shape);
        Assert.Equal(highest, sheet.Cells[(1, 0)].Icon!.Shape);
    }

    [Fact]
    public void Level_icons_step_through_their_levels()
    {
        var sheet = Read(Column("A", 0, 20, 40, 60, 80), ConditionalFormatting("A1:A5", IconSet(1, "5Quarters", Thresholds(0, 20, 40, 60, 80))));
        Assert.Equal([0, 1, 2, 3, 4], Enumerable.Range(0, 5).Select(i => sheet.Cells[(i, 0)].Icon!.Level));
    }

    [Fact]
    public void Icon_sets_support_reverse_hidden_values_numbers_and_gte()
    {
        var values = Column("A", 1, 4, 8);
        var reversed = Read(values, ConditionalFormatting("A1:A3", IconSet(1, "3Arrows", Thresholds(0, 33, 67), "reverse=\"1\"")));
        Assert.Equal(IconShape.ArrowUp, reversed.Cells[(0, 0)].Icon!.Shape);
        Assert.Equal(IconShape.ArrowDown, reversed.Cells[(2, 0)].Icon!.Shape);
        var hidden = Read(values, ConditionalFormatting("A1:A3", IconSet(1, "3Arrows", Thresholds(0, 33, 67), "showValue=\"0\"")));
        Assert.Equal("", hidden.Cells[(0, 0)].Text);
        Assert.NotNull(hidden.Cells[(0, 0)].Icon);
        var numbers = Column("A", 5, 6, 7);
        var strict = Read(numbers, ConditionalFormatting("A1:A3", IconSet(1, "3Arrows", Cfvo("percent", "0") + Cfvo("num", "6").Replace("/>", " gte=\"0\"/>") + Cfvo("num", "7"))));
        Assert.Equal(new[] { IconShape.ArrowDown, IconShape.ArrowDown, IconShape.ArrowUp }, Enumerable.Range(0, 3).Select(i => strict.Cells[(i, 0)].Icon!.Shape));
    }

    [Fact]
    public void Scale_bar_and_icon_rules_do_not_apply_to_text_cells()
    {
        var sheet = Read(Row(1, Inline("A1", "x")), ConditionalFormatting("A1", IconSet(1, "3Arrows", Thresholds(0, 33, 67)), DataBar(2)));
        Assert.Null(sheet.Cells[(0, 0)].Icon);
        Assert.Null(sheet.Cells[(0, 0)].DataBar);
    }

    [Fact]
    public void A_stop_if_true_flag_is_ignored_by_scale_bar_and_icon_rules_but_a_matching_formula_rule_stops_them()
    {
        var iconRule = IconSet(2, "3Arrows", Thresholds(0, 33, 67)).Replace("priority=\"2\"", "priority=\"2\" stopIfTrue=\"1\"");
        var sheet = Read(Column("A", 1, 9), ConditionalFormatting("A1:A2", iconRule, DataBar(3)));
        Assert.NotNull(sheet.Cells[(0, 0)].Icon);
        Assert.NotNull(sheet.Cells[(0, 0)].DataBar);
        var stopped = Read(Column("A", 1, 9), ConditionalFormatting("A1:A2", CellIs("lessThan", "5", 1, 0, stopIfTrue: true), iconRule));
        Assert.Null(stopped.Cells[(0, 0)].Icon);
        Assert.NotNull(stopped.Cells[(1, 0)].Icon);
    }

    [Fact]
    public void Dxf_font_border_and_number_format_are_baked_into_the_cell()
    {
        var font = "<b/><i/><u/><strike/><color rgb=\"FFFF0000\"/>";
        var border = "<left style=\"thin\"><color rgb=\"FF0000FF\"/></left>";
        var dxfs = new[] { Dxf(font: font, border: border, numFmt: "<numFmt numFmtId=\"170\" formatCode=\"0.0%\"/>") };
        var sheet = Read(Column("A", 0.256), ConditionalFormatting("A1", CellIs("greaterThan", "0", 1, 0)), dxfs);
        var cell = sheet.Cells[(0, 0)];
        Assert.True(cell.Style.Font.Bold);
        Assert.True(cell.Style.Font.Italic);
        Assert.True(cell.Style.Font.Strike);
        Assert.Equal(UnderlineStyle.Single, cell.Style.Font.Underline);
        Assert.Equal(Hex("FF0000"), cell.Style.Font.Color);
        Assert.Equal(BorderStyle.Single, cell.Style.Borders.Left.Style);
        Assert.Equal(Hex(Blue), cell.Style.Borders.Left.Color);
        Assert.False(cell.Style.Borders.Right.IsVisible);
        Assert.Equal("25.6%", cell.Text);
        Assert.Equal(0.256, cell.Value);
    }

    [Fact]
    public void A_dxf_can_switch_bold_off()
    {
        var builder = new XlsxBuilder
        {
            Styles = "<fonts><font><b/><name val=\"Calibri\"/></font></fonts><fills><fill/></fills><borders><border/></borders><cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>" +
                     $"<dxfs count=\"1\">{Dxf(font: "<b val=\"0\"/>")}</dxfs>",
        }.AddSheet("S", Row(1, C("A1", "5", style: 0)), tail: ConditionalFormatting("A1", CellIs("greaterThan", "0", 1, 0)));
        var cell = XlsxReader.Read(builder.BuildStream()).Sheets[0].Cells[(0, 0)];
        Assert.False(cell.Style.Font.Bold);
    }

    [Fact]
    public void Dxf_fill_replaces_a_gradient_and_rich_text_runs_follow_the_dxf_font()
    {
        var builder = new XlsxBuilder
        {
            SharedStrings = "<si><r><rPr><sz val=\"11\"/></rPr><t>ab</t></r><r><rPr><i/></rPr><t>cd</t></r></si>",
            Styles = StylesWithDxfs(Dxf(Red, font: "<b/>")),
        }.AddSheet("S", Row(1, S("A1", 0, style: 0)), tail: ConditionalFormatting("A1", Expression("LEN(A1)=4", 1, 0)));
        var cell = XlsxReader.Read(builder.BuildStream()).Sheets[0].Cells[(0, 0)];
        Assert.All(cell.Runs!, run => Assert.True(run.Font.Bold));
        Assert.True(cell.Runs![1].Font.Italic);
        Assert.Equal(Hex(Red), cell.Style.Fill);
    }

    [Fact]
    public void Supported_rules_do_not_warn_but_unsupported_ones_warn_once_per_sheet()
    {
        var warnings = new ConversionWarningList();
        Read(Column("A", 1, 2), ConditionalFormatting("A1:A2", CellIs("greaterThan", "1", 1, 0), DataBar(2)), warnings: warnings);
        Assert.Empty(warnings.Warnings);

        var tail = ConditionalFormatting("A1", CfRule("unknownKind", 1, "", "", 0)) +
                   ConditionalFormatting("A2", IconSet(2, "7Faces", Thresholds(0, 20, 40, 60, 80))) +
                   ConditionalFormatting("A1:A2", Expression("SUMPRODUCT(A1:A2)>1", 3, 0), Expression("Other!A1>1", 4, 0), CellIs("greaterThan", "1", 5, 0));
        var sheet = Read(Column("A", 1, 2), tail, warnings: warnings);
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("xlsx.conditional-formatting", warning.Code);
        Assert.Contains("unknownKind", warning.Message);
        Assert.Contains("7Faces", warning.Message);
        Assert.Contains("expression", warning.Message);
        // The supported rule still applies.
        Assert.Equal("A2", FilledRows(sheet));
    }

    private static string X14Block(string sqref, string rule) =>
        "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"" + X14Ns + "\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
        $"<x14:conditionalFormattings><x14:conditionalFormatting>{rule}<xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";

    private static string X14Cfvo(string type, string? value = null) =>
        value is null ? $"<x14:cfvo type=\"{type}\"/>" : $"<x14:cfvo type=\"{type}\"><xm:f>{value}</xm:f></x14:cfvo>";

    [Fact]
    public void X14_only_icon_sets_are_read_from_the_extension_list()
    {
        var warnings = new ConversionWarningList();
        var rule = "<x14:cfRule type=\"iconSet\" priority=\"1\" id=\"{00000000-0000-0000-0000-000000000009}\"><x14:iconSet iconSet=\"3Stars\">" +
                   X14Cfvo("percent", "0") + X14Cfvo("percent", "33") + X14Cfvo("percent", "67") + "</x14:iconSet></x14:cfRule>";
        var sheet = Read(Column("A", 1, 5, 10), X14Block("A1:A3", rule), warnings: warnings);
        Assert.Empty(warnings.Warnings);
        Assert.Equal([0, 1, 2], Enumerable.Range(0, 3).Select(i => sheet.Cells[(i, 0)].Icon!.Level));
        Assert.All(Enumerable.Range(0, 3), i => Assert.Equal(IconShape.Star, sheet.Cells[(i, 0)].Icon!.Shape));
    }

    [Fact]
    public void X14_only_rules_use_their_inline_dxf_and_xm_formulas()
    {
        var rule = "<x14:cfRule type=\"cellIs\" priority=\"1\" operator=\"greaterThan\" id=\"{00000000-0000-0000-0000-00000000000A}\"><xm:f>4</xm:f>" +
                   "<x14:dxf><fill><patternFill><bgColor rgb=\"FF00FF00\"/></patternFill></fill></x14:dxf></x14:cfRule>";
        var sheet = Read(Column("A", 1, 10, 5), X14Block("A1:A3", rule));
        Assert.Equal("A2,A3", FilledRows(sheet, "00FF00"));
    }

    [Fact]
    public void X14_only_rules_with_unsupported_content_warn()
    {
        var warnings = new ConversionWarningList();
        var rule = "<x14:cfRule type=\"expression\" priority=\"1\" id=\"{00000000-0000-0000-0000-00000000000B}\"><xm:f>SUMPRODUCT(A1:A2)&gt;1</xm:f><x14:dxf><font><b/></font></x14:dxf></x14:cfRule>";
        Read(Column("A", 1), X14Block("A1", rule), warnings: warnings);
        Assert.Contains("expression", Assert.Single(warnings.Warnings).Message);
    }

    [Fact]
    public void A_data_bar_extension_is_not_read_as_a_rule_of_its_own()
    {
        var warnings = new ConversionWarningList();
        var tail = ConditionalFormatting("A1:A2", DataBar(1, extensionId: "{00000000-0000-0000-0000-00000000000C}")) +
                   X14Block("A1:A2", "<x14:cfRule type=\"dataBar\" id=\"{00000000-0000-0000-0000-00000000000C}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\"/></x14:cfRule>");
        var sheet = Read(Column("A", 1, 2), tail, warnings: warnings);
        Assert.Empty(warnings.Warnings);
        Assert.NotNull(sheet.Cells[(1, 0)].DataBar);
    }

    [Fact]
    public void A_rule_without_a_valid_dxf_matches_but_changes_nothing()
    {
        var sheet = Read(Column("A", 5), ConditionalFormatting("A1", CfRule("cellIs", 1, "operator=\"greaterThan\"", "<formula>1</formula>", 99)));
        Assert.Null(sheet.Cells[(0, 0)].Style.Fill);
    }

    [Theory]
    [InlineData("today", "A3")]
    [InlineData("yesterday", "A2")]
    [InlineData("tomorrow", "A4")]
    [InlineData("last7Days", "A1,A2,A3,A6,A9")]           // May 9 .. May 15
    [InlineData("thisWeek", "A2,A3,A4,A5")]               // Sunday May 12 .. Saturday May 18
    [InlineData("lastWeek", "A1,A6,A9")]                  // Sunday May 5 .. Saturday May 11
    [InlineData("nextWeek", "A7")]                        // May 19 .. May 25
    [InlineData("thisMonth", "A1,A2,A3,A4,A5,A6,A7,A9")]
    [InlineData("lastMonth", "A8")]
    [InlineData("nextMonth", "A10")]
    public void Time_period_rules_are_relative_to_today(string period, string expected)
    {
        // Today is Wednesday 2024-05-15. Rows: 1 = May 10, 2 = May 14, 3 = May 15, 4 = May 16, 5 = May 18 (Sat), 6 = May 11 (Sat),
        // 7 = May 19 (Sun), 8 = April 30, 9 = May 9, 10 = June 1.
        static double Serial(int month, int day) => (new DateTime(2024, month, day) - new DateTime(1899, 12, 30)).TotalDays;
        var sheet = Read(Column("A", Serial(5, 10), Serial(5, 14), Serial(5, 15), Serial(5, 16), Serial(5, 18), Serial(5, 11), Serial(5, 19), Serial(4, 30), Serial(5, 9), Serial(6, 1)),
            ConditionalFormatting("A1:A10", CfRule("timePeriod", 1, $" timePeriod=\"{period}\"", dxfId: 0)), today: new DateTime(2024, 5, 15));
        Assert.Equal(expected, FilledRows(sheet, Red));
    }

    [Fact]
    public void Time_period_rules_ignore_text_and_blanks()
    {
        var sheet = Read(Row(1, Inline("A1", "today")), ConditionalFormatting("A1:A2", CfRule("timePeriod", 1, " timePeriod=\"today\"", dxfId: 0)), today: new DateTime(2024, 5, 15));
        Assert.Empty(FilledRows(sheet));
    }

    private static (Worksheet Main, Worksheet Other, ConversionWarningList Warnings) ReadTwo(string rows, string tail, string otherRows, string? names = null, string otherName = "Lists")
    {
        var warnings = new ConversionWarningList();
        var builder = new XlsxBuilder { Styles = StylesWithDxfs([Dxf(Red), Dxf(Green)]), DefinedNames = names }.AddSheet("S", rows, tail: tail).AddSheet(otherName, otherRows);
        var workbook = XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });
        return (workbook.Sheets[0], workbook.Sheets[1], warnings);
    }

    [Fact]
    public void Formulas_can_read_cells_of_other_sheets_even_when_that_sheet_comes_later()
    {
        // A cell is filled when it exceeds the limit stored on the (later) sheet "Lists"; the relative reference moves with the cell.
        var rows = Column("A", 1, 10, 5);
        var other = Row(1, C("A1", "4")) + Row(2, C("A2", "20")) + Row(3, C("A3", "1"));
        var (main, _, warnings) = ReadTwo(rows, ConditionalFormatting("A1:A3", Expression("A1>Lists!A1", 1, 0)), other);
        Assert.Empty(warnings.Warnings);
        Assert.Equal("A3", FilledRows(main)); // 1>4 no, 10>20 no, 5>1 yes
    }

    [Fact]
    public void Quoted_sheet_names_absolute_ranges_and_functions_over_other_sheets_work()
    {
        var rows = Column("A", 1, 10, 5);
        var other = Row(1, C("A1", "1")) + Row(2, C("A2", "10"));
        var rule = Expression("COUNTIF('My list'!$A$1:$A$3,A1)>0", 1, 0);
        var (main, _, warnings) = ReadTwo(rows, ConditionalFormatting("A1:A3", rule), other, otherName: "My list");
        Assert.Empty(warnings.Warnings);
        Assert.Equal("A1,A2", FilledRows(main)); // 1 and 10 are in the list, 5 is not
        var sum = ReadTwo(rows, ConditionalFormatting("A1:A3", Expression("A1<SUM('My list'!$A$1:$A$3)/2", 1, 0)), other, otherName: "My list").Main;
        Assert.Equal("A1,A3", FilledRows(sum)); // 11 / 2 = 5.5: 1 and 5 are below it
    }

    [Fact]
    public void Cell_is_rules_and_thresholds_can_refer_to_other_sheets()
    {
        var rows = Column("A", 1, 10, 5);
        var other = Row(1, C("A1", "4"));
        var (main, _, warnings) = ReadTwo(rows, ConditionalFormatting("A1:A3", CellIs("greaterThan", "Lists!$A$1", 1, 0)), other);
        Assert.Empty(warnings.Warnings);
        Assert.Equal("A2,A3", FilledRows(main));
    }

    [Fact]
    public void Defined_names_standing_for_references_ranges_and_constants_are_resolved()
    {
        var rows = Column("A", 1, 10, 5);
        var other = Row(1, C("A1", "4")) + Row(2, C("A2", "6"));
        var names = "<definedName name=\"Limit\">Lists!$A$1</definedName><definedName name=\"Bad\" localSheetId=\"0\">Lists!$A$2:$A$2</definedName><definedName name=\"Fixed\">7</definedName>";
        var (limit, _, w1) = ReadTwo(rows, ConditionalFormatting("A1:A3", Expression("A1>Limit", 1, 0)), other, names);
        Assert.Empty(w1.Warnings);
        Assert.Equal("A2,A3", FilledRows(limit));
        var (fixedValue, _, w2) = ReadTwo(rows, ConditionalFormatting("A1:A3", Expression("A1>Fixed", 1, 0)), other, names);
        Assert.Empty(w2.Warnings);
        Assert.Equal("A2", FilledRows(fixedValue));
        var (range, _, w3) = ReadTwo(rows, ConditionalFormatting("A1:A3", Expression("A1>MAX(Bad)", 1, 0)), other, names);
        Assert.Empty(w3.Warnings);
        Assert.Equal("A2", FilledRows(range)); // a sheet-scoped name: > 6
    }

    [Fact]
    public void Names_that_are_not_plain_references_and_unknown_names_are_reported_as_unsupported()
    {
        var names = "<definedName name=\"Dynamic\">OFFSET(Lists!$A$1,0,0,2,1)</definedName>";
        var (main, _, warnings) = ReadTwo(Column("A", 1, 10), ConditionalFormatting("A1:A2", Expression("A1>MAX(Dynamic)", 1, 0), Expression("A1>Nothing", 2, 1)), Row(1, C("A1", "4")), names);
        Assert.Equal("", FilledRows(main));
        Assert.Contains("expression", Assert.Single(warnings.Warnings).Message);
    }

    [Fact]
    public void A_rule_reading_a_sheet_that_does_not_exist_never_applies_and_says_so()
    {
        var (main, _, warnings) = ReadTwo(Column("A", 1, 10), ConditionalFormatting("A1:A2", Expression("A1>Gone!A1", 1, 0)), Row(1, C("A1", "4")));
        Assert.Equal("", FilledRows(main));
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("xlsx.conditional-formatting", warning.Code);
        Assert.Contains("'Gone'", warning.Message);
    }

    private static string ThreeThresholds => Thresholds(0, 33, 67);

    [Fact]
    public void Custom_icons_replace_single_levels_with_icons_of_any_set_or_none()
    {
        const string id = "{55555555-5555-5555-5555-555555555555}";
        var tail = ConditionalFormatting("A1:A3", IconSet(1, "3Arrows", ThreeThresholds, extensionId: id)) +
                   IconSetExtension(id, "3Arrows", ("3Flags", 0), ("NoIcons", 0), ("3Stars", 2));
        var sheet = Read(Column("A", 1, 50, 100), tail);
        Assert.Equal(new CellIcon(IconShape.Flag, Hex("E03030")), sheet.Cells[(0, 0)].Icon);
        Assert.Null(sheet.Cells[(1, 0)].Icon);
        Assert.Equal(IconShape.Star, sheet.Cells[(2, 0)].Icon!.Shape);
        Assert.Equal(2, sheet.Cells[(2, 0)].Icon!.Level);
        // Without the extension the set is used as it is.
        var plain = Read(Column("A", 1, 50, 100), ConditionalFormatting("A1:A3", IconSet(1, "3Arrows", ThreeThresholds)));
        Assert.Equal([IconShape.ArrowDown, IconShape.ArrowSideways, IconShape.ArrowUp], Enumerable.Range(0, 3).Select(i => plain.Cells[(i, 0)].Icon!.Shape));
    }

    [Fact]
    public void Custom_icons_of_x14_only_rules_and_invalid_choices_fall_back_to_the_sets_icons()
    {
        var rule = "<x14:cfRule type=\"iconSet\" priority=\"1\" id=\"{00000000-0000-0000-0000-0000000000AA}\"><x14:iconSet iconSet=\"3Triangles\" custom=\"1\">" +
                   X14Cfvo("percent", "0") + X14Cfvo("percent", "33") + X14Cfvo("percent", "67") +
                   "<x14:cfIcon iconSet=\"3Symbols\" iconId=\"2\"/><x14:cfIcon iconSet=\"3Symbols\" iconId=\"9\"/></x14:iconSet></x14:cfRule>";
        var sheet = Read(Column("A", 1, 50, 100), X14Block("A1:A3", rule));
        Assert.Equal(IconShape.Check, sheet.Cells[(0, 0)].Icon!.Shape);
        Assert.Equal(IconShape.Dash, sheet.Cells[(1, 0)].Icon!.Shape); // an unknown icon id keeps the set's own
        Assert.Equal(IconShape.TriangleUp, sheet.Cells[(2, 0)].Icon!.Shape); // no choice for the third level
    }

    private static string AxisExtension(string id, string attributes, string children = "") => DataBarExtension(id, gradient: false, attributes, children);

    [Fact]
    public void A_data_bar_axis_in_the_middle_splits_the_cell_in_half()
    {
        const string id = "{66666666-6666-6666-6666-666666666666}";
        var tail = ConditionalFormatting("A1:A3", DataBar(1, extensionId: id)) + AxisExtension(id, "axisPosition=\"middle\"", "<x14:axisColor rgb=\"FFFF0000\"/>");
        var sheet = Read(Column("A", -50, 0, 100), tail);
        var negative = sheet.Cells[(0, 0)].DataBar!;
        Assert.Equal((0.0, 0.5, 0.5), (negative.Start, negative.End, negative.Axis!.Value)); // the minimum fills the left half
        var positive = sheet.Cells[(2, 0)].DataBar!;
        Assert.Equal((0.5, 1.0), (positive.Start, positive.End));
        Assert.Equal(Hex("FF0000"), positive.AxisColor);
        Assert.Equal(0.75, Read(Column("A", -50, 50, 100), tail).Cells[(1, 0)].DataBar!.End); // half of the maximum
    }

    [Fact]
    public void An_automatic_axis_is_drawn_in_black_by_default_and_none_draws_no_axis()
    {
        const string id = "{77777777-7777-7777-7777-777777777777}";
        var automatic = Read(Column("A", -50, 0, 100), ConditionalFormatting("A1:A3", DataBar(1, extensionId: id)) + AxisExtension(id, ""));
        Assert.Equal(1.0 / 3, automatic.Cells[(2, 0)].DataBar!.Axis!.Value, 6);
        Assert.Equal(Color.Black, automatic.Cells[(2, 0)].DataBar!.AxisColor);

        var none = Read(Column("A", -50, 0, 100), ConditionalFormatting("A1:A3", DataBar(1, extensionId: id)) + AxisExtension(id, "axisPosition=\"none\"", "<x14:negativeFillColor rgb=\"FF00FF00\"/>"));
        var low = none.Cells[(0, 0)].DataBar!;
        Assert.Null(low.Axis);
        Assert.False(low.Negative);
        Assert.Equal(Hex("00FF00"), low.Color); // negative values keep their own color
        Assert.Equal(0.0, low.Start);
        Assert.True(none.Cells[(0, 0)].DataBar!.End < none.Cells[(2, 0)].DataBar!.End); // bars grow with the value from the left
    }

    [Fact]
    public void A_right_to_left_data_bar_is_mirrored()
    {
        const string id = "{88888888-8888-8888-8888-888888888888}";
        var tail = ConditionalFormatting("A1:A3", DataBar(1, extensionId: id)) + AxisExtension(id, "direction=\"rightToLeft\"");
        var sheet = Read(Column("A", 1, 50, 100), tail);
        var full = sheet.Cells[(2, 0)].DataBar!;
        Assert.True(full.Mirrored);
        Assert.Equal((0.0, 1.0), (full.Start, full.End));
        var small = sheet.Cells[(0, 0)].DataBar!;
        Assert.Equal(1.0, small.End, 6); // the base is at the right edge
        Assert.True(small.Start > 0.5);
    }
}
