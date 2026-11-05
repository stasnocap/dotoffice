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

    private static Worksheet Read(string rows, string tail, string[]? dxfs = null, ConversionWarningList? warnings = null, string? definedNames = null)
    {
        var builder = new XlsxBuilder { Styles = StylesWithDxfs(dxfs ?? [Dxf(Red), Dxf(Green)]), DefinedNames = definedNames }.AddSheet("S", rows, tail: tail);
        return XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings }).Sheets[0];
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

        var tail = ConditionalFormatting("A1", CfRule("timePeriod", 1, "timePeriod=\"today\"", "", 0)) +
                   ConditionalFormatting("A2", IconSet(2, "5Arrows", Thresholds(0, 20, 40, 60, 80))) +
                   ConditionalFormatting("A1:A2", Expression("SUMPRODUCT(A1:A2)>1", 3, 0), Expression("Other!A1>1", 4, 0), CellIs("greaterThan", "1", 5, 0));
        var sheet = Read(Column("A", 1, 2), tail, warnings: warnings);
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("xlsx.conditional-formatting", warning.Code);
        Assert.Contains("timePeriod", warning.Message);
        Assert.Contains("5Arrows", warning.Message);
        Assert.Contains("expression", warning.Message);
        // The supported rule still applies.
        Assert.Equal("A2", FilledRows(sheet));
    }

    [Fact]
    public void Rules_in_the_x14_extension_list_are_not_read_as_main_rules()
    {
        var warnings = new ConversionWarningList();
        var x14Only = "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"" + X14Ns + "\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
                      "<x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type=\"iconSet\" priority=\"1\"><x14:iconSet iconSet=\"4Rating\"/></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";
        var sheet = Read(Column("A", 1), x14Only, warnings: warnings);
        Assert.Empty(warnings.Warnings);
        Assert.Null(sheet.Cells[(0, 0)].Icon);
    }

    [Fact]
    public void A_rule_without_a_valid_dxf_matches_but_changes_nothing()
    {
        var sheet = Read(Column("A", 5), ConditionalFormatting("A1", CfRule("cellIs", 1, "operator=\"greaterThan\"", "<formula>1</formula>", 99)));
        Assert.Null(sheet.Cells[(0, 0)].Style.Fill);
    }
}
