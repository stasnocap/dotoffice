using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;

namespace DotOffice.Samples;

/// <summary>
/// Generates a self-contained showcase XLSX (raw SpreadsheetML) exercising the XLSX → PDF converter: number and date formats,
/// styles with fills and borders, merged cells, print titles and manual breaks, fit-to-width, wrapping, overflow, alignment,
/// indentation, headers and footers, and a hidden sheet whose text must not be printed. Everything in it is generated.
/// </summary>
internal static class ShowcaseXlsx
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    // cellXfs indexes (see Styles)
    private const int Default = 0, Title = 1, Header = 2, Money = 3, Percent = 4, IsoDate = 5, Time = 6, Accounting = 7, Wrap = 8,
        Right = 9, Center = 10, Indent = 11, Boxed = 12, Scientific = 13, Percent2 = 14, LongDate = 15, Bold = 16,
        Rot45 = 17, Rot90 = 18, RotDown = 19, Stacked = 20, RotWrap = 21, GradLinear = 22, GradDiagonal = 23, GradVertical = 24, GradCenter = 25, GradCorner = 26,
        Fraction1 = 27, Fraction2 = 28, Eighths = 29, IndexedRed = 30, FillAlign = 31;

    public const string HiddenMarker = "HIDDENSHEETMARKER";

    public static void Write(string path)
    {
        var sheets = new List<(string Name, string Xml, string State)>
        {
            ("Formats", Formats(), "visible"),
            ("Long table", LongTable(), "visible"),
            ("Wide", Wide(), "visible"),
            ("Text", TextSheet(), "visible"),
            ("Rich text", RichTextSheet(), "visible"),
            ("Rotation", RotationSheet(), "visible"),
            ("Gradients", GradientSheet(), "visible"),
            ("Pictures", PicturesSheet(), "visible"),
            ("Shapes", ShapesSheet(), "visible"),
            ("Shape paths", ShapePathsSheet(), "visible"),
            ("Shape presets", ShapePresetsSheet(), "visible"),
            ("Shape fills", ShapeFillsSheet(), "visible"),
            ("Charts", ChartsSheet(), "visible"),
            ("More charts", MoreChartsSheet(), "visible"),
            ("Chart options", ChartOptionsSheet(), "visible"),
            ("Chart extras", ChartExtrasSheet(), "visible"),
            ("Chart text", ChartTextSheet(), "visible"),
            ("Conditional", ConditionalSheet(), "visible"),
            ("Conditional links", ConditionalLinksSheet(), "visible"),
            ("Conditional lookups", ConditionalLookupsSheet(), "visible"),
            ("Pivot", PivotSheet("Pivot table cached cells (no pivot style: nothing to warn about)"), "visible"),
            ("Pivot styled", PivotSheet("Pivot table with a pivot style and a pivot-area format on the corner (approximated, reported)"), "visible"),
            ("Pivot banded", PivotSheet("Pivot table with a banded Medium style and formats on the East data (pink) and the North label (green) (approximated, reported)"), "visible"),
            ("Pivot dark", PivotSheet("Pivot table with a banded Dark style (approximated, reported)"), "visible"),
            ("Pivot custom", PivotSheet("Pivot table with the workbook's own table style from styles.xml (header, two stripes, total row; reported)"), "visible"),
            ("Conditional tables", ConditionalTablesSheet(), "visible"),
            ("Span A", Sheet("", Row(1, Num("H1", 10))), "hidden"),
            ("Span B", Sheet("", Row(1, Num("H1", 25))), "hidden"),
            ("Pivot fields", PivotFieldsSheet(), "visible"),
            ("Paper", PaperSheet(), "visible"),
            ("Areas", AreasSheet(), "visible"),
            ("Comments", CommentsSheet(), "visible"),
            ("Titles", TitlesSheet(), "visible"),
            ("Numbering", NumberingSheet(), "visible"),
            ("Secret", Hidden(), "hidden"),
            ("Limits", Limits(), "hidden"),
        };

        using var stream = File.Create(path);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create);
        var content = new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        var rels = new StringBuilder();
        var sheetEntries = new StringBuilder();
        for (int i = 0; i < sheets.Count; i++)
        {
            var (name, xml, state) = sheets[i];
            Entry(zip, $"xl/worksheets/sheet{i + 1}.xml", xml);
            content.Append($"<Override PartName=\"/xl/worksheets/sheet{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            rels.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"{RelType}worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
            var stateAttr = state == "visible" ? "" : $" state=\"{state}\"";
            sheetEntries.Append($"<sheet name=\"{SecurityElement.Escape(name)}\" sheetId=\"{i + 1}\"{stateAttr} r:id=\"rId{i + 1}\"/>");
        }
        rels.Append($"<Relationship Id=\"rId{sheets.Count + 1}\" Type=\"{RelType}styles\" Target=\"styles.xml\"/>");
        WritePictures(zip, sheets.FindIndex(x => x.Name == "Pictures") + 1, content);
        WriteShapes(zip, sheets.FindIndex(x => x.Name == "Shapes") + 1);
        WritePivot(zip, sheets.FindIndex(x => x.Name == "Pivot") + 1, "");
        WritePivot(zip, sheets.FindIndex(x => x.Name == "Pivot styled") + 1, "PivotStyleLight16",
            "<formats count=\"1\"><format dxfId=\"5\"><pivotArea type=\"origin\" outline=\"0\" fieldPosition=\"0\"/></format></formats>");
        WritePivot(zip, sheets.FindIndex(x => x.Name == "Pivot banded") + 1, "PivotStyleMedium2",
            "<formats count=\"2\"><format dxfId=\"0\"><pivotArea field=\"0\" type=\"normal\"><references count=\"1\"><reference field=\"0\" count=\"1\"><x v=\"2\"/></reference></references></pivotArea></format>" +
            "<format dxfId=\"1\"><pivotArea dataOnly=\"0\" labelOnly=\"1\" field=\"0\" type=\"normal\"><references count=\"1\"><reference field=\"0\" count=\"1\"><x v=\"0\"/></reference></references></pivotArea></format></formats>", stripes: true);
        WritePivot(zip, sheets.FindIndex(x => x.Name == "Pivot dark") + 1, "PivotStyleDark3", "", stripes: true);
        WritePivot(zip, sheets.FindIndex(x => x.Name == "Pivot custom") + 1, "Showcase pivot", "", stripes: true);
        WritePivotFields(zip, sheets.FindIndex(x => x.Name == "Pivot fields") + 1);
        WriteShapePaths(zip, sheets.FindIndex(x => x.Name == "Shape paths") + 1);
        WriteShapePresets(zip, sheets.FindIndex(x => x.Name == "Shape presets") + 1);
        WriteShapeFills(zip, sheets.FindIndex(x => x.Name == "Shape fills") + 1);
        WriteCharts(zip, sheets.FindIndex(x => x.Name == "Charts") + 1);
        WriteMoreCharts(zip, sheets.FindIndex(x => x.Name == "More charts") + 1);
        WriteChartOptions(zip, sheets.FindIndex(x => x.Name == "Chart options") + 1);
        WriteChartExtras(zip, sheets.FindIndex(x => x.Name == "Chart extras") + 1);
        WriteChartText(zip, sheets.FindIndex(x => x.Name == "Chart text") + 1);
        WriteComments(zip, sheets.FindIndex(x => x.Name == "Comments") + 1, content);
        WriteConditionalTables(zip, sheets.FindIndex(x => x.Name == "Conditional tables") + 1);
        content.Append("</Types>");

        Entry(zip, "[Content_Types].xml", content.ToString());
        Entry(zip, "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"{RelType}officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/></Relationships>");
        Entry(zip, "docProps/core.xml",
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<dc:title>DotOffice XLSX showcase</dc:title><dc:creator>DotOffice</dc:creator></cp:coreProperties>");
        Entry(zip, "xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");
        Entry(zip, "xl/workbook.xml",
            $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets>{sheetEntries}</sheets>" +
            "<definedNames><definedName name=\"Limit\">Limits!$A$1</definedName><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"1\">'Long table'!$1:$2</definedName>" +
            $"<definedName name=\"_xlnm.Print_Area\" localSheetId=\"{sheets.FindIndex(x => x.Name == "Areas")}\">Areas!$A$1:$C$3,Areas!$E$5:$F$7,Areas!$A$10:$B$11</definedName>" +
            $"<definedName name=\"_xlnm.Print_Area\" localSheetId=\"{sheets.FindIndex(x => x.Name == "Titles")}\">Titles!$A$5:$B$124</definedName>" +
            $"<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"{sheets.FindIndex(x => x.Name == "Titles")}\">Titles!$1:$1</definedName></definedNames></workbook>");
        Entry(zip, "xl/styles.xml", Styles);
    }

    // ------------------------------------------------------------------ styles

    private const string Styles =
        $"<styleSheet xmlns=\"{Main}\">" +
        "<numFmts count=\"7\"><numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/><numFmt numFmtId=\"165\" formatCode=\"#,##0;[Red]\\(#,##0\\)\"/>" +
        "<numFmt numFmtId=\"166\" formatCode=\"dddd, mmmm d, yyyy\"/>" +
        "<numFmt numFmtId=\"168\" formatCode=\"# ?/?\"/><numFmt numFmtId=\"169\" formatCode=\"# ??/??\"/><numFmt numFmtId=\"170\" formatCode=\"# ?/8\"/><numFmt numFmtId=\"171\" formatCode=\"[Color3]0.0\"/></numFmts>" +
        "<fonts count=\"4\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
        "<font><b/><sz val=\"16\"/><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/></font><font><i/><sz val=\"11\"/><color rgb=\"FF7F7F7F\"/><name val=\"Calibri\"/></font></fonts>" +
        "<fills count=\"9\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF1F3864\"/></patternFill></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFD9E1F2\"/></patternFill></fill>" +
        "<fill><gradientFill degree=\"0\"><stop position=\"0\"><color rgb=\"FFFFFFFF\"/></stop><stop position=\"1\"><color rgb=\"FF4472C4\"/></stop></gradientFill></fill>" +           // 4 left to right
        "<fill><gradientFill degree=\"45\"><stop position=\"0\"><color theme=\"5\"/></stop><stop position=\"0.5\"><color rgb=\"FFFFFF00\"/></stop><stop position=\"1\"><color indexed=\"4\"/></stop></gradientFill></fill>" + // 5 diagonal, 3 stops
        "<fill><gradientFill degree=\"90\"><stop position=\"0\"><color rgb=\"FF70AD47\"/></stop><stop position=\"1\"><color rgb=\"FF1E4620\"/></stop></gradientFill></fill>" +             // 6 top to bottom
        "<fill><gradientFill type=\"path\" left=\"0.5\" right=\"0.5\" top=\"0.5\" bottom=\"0.5\"><stop position=\"0\"><color rgb=\"FFFFFFFF\"/></stop><stop position=\"1\"><color rgb=\"FFC00000\"/></stop></gradientFill></fill>" + // 7 from center
        "<fill><gradientFill type=\"path\"><stop position=\"0\"><color rgb=\"FFFFC000\"/></stop><stop position=\"1\"><color rgb=\"FF7030A0\"/></stop></gradientFill></fill></fills>" +
        "<borders count=\"3\"><border><left/><right/><top/><bottom/><diagonal/></border>" +
        "<border><left/><right/><top/><bottom style=\"medium\"><color rgb=\"FF1F3864\"/></bottom><diagonal/></border>" +
        "<border><left style=\"thin\"><color auto=\"1\"/></left><right style=\"thin\"><color auto=\"1\"/></right><top style=\"thin\"><color auto=\"1\"/></top><bottom style=\"thin\"><color auto=\"1\"/></bottom><diagonal/></border></borders>" +
        "<cellXfs count=\"32\">" +
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                       // 0 default
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"2\" borderId=\"0\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +                                // 1 title
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"1\"><alignment horizontal=\"center\"/></xf>" +                                                   // 2 header
        "<xf numFmtId=\"4\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                       // 3 #,##0.00
        "<xf numFmtId=\"9\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                        // 4 0%
        "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 5 yyyy-mm-dd
        "<xf numFmtId=\"18\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                       // 6 h:mm AM/PM
        "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 7 accounting-ish
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"2\"><alignment vertical=\"top\" wrapText=\"1\"/></xf>" +                                          // 8 wrap
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment horizontal=\"right\"/></xf>" +                                                   // 9 right
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment horizontal=\"center\"/></xf>" +                                                  // 10 center
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment horizontal=\"left\" indent=\"2\"/></xf>" +                                        // 11 indent
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"2\"/>" +                                                                                        // 12 boxed
        "<xf numFmtId=\"11\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                       // 13 0.00E+00
        "<xf numFmtId=\"10\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                       // 14 0.00%
        "<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 15 long date
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\"/>" +                                                                                        // 16 bold
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"2\"><alignment horizontal=\"center\" textRotation=\"45\"/></xf>" +                                  // 17 rotated 45°
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"2\"><alignment horizontal=\"center\" vertical=\"center\" textRotation=\"90\"/></xf>" +                  // 18 rotated 90°
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"2\"><alignment horizontal=\"center\" vertical=\"center\" textRotation=\"180\"/></xf>" +                 // 19 reading down
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"2\"><alignment horizontal=\"center\" vertical=\"center\" textRotation=\"255\"/></xf>" +                 // 20 stacked
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"2\"><alignment vertical=\"center\" wrapText=\"1\" textRotation=\"30\"/></xf>" +                      // 21 rotated and wrapped
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"4\" borderId=\"2\"/>" +                                                                                         // 22 gradient left to right
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"5\" borderId=\"2\"/>" +                                                                                         // 23 diagonal
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"6\" borderId=\"2\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +                                // 24 vertical
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"7\" borderId=\"2\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +                                // 25 from center
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"8\" borderId=\"2\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +                                // 26 from corner
        "<xf numFmtId=\"168\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 27 # ?/?
        "<xf numFmtId=\"169\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 28 # ??/??
        "<xf numFmtId=\"170\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 29 # ?/8
        "<xf numFmtId=\"171\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +                                                                                      // 30 [Color3]0.0
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"2\"><alignment horizontal=\"fill\"/></xf>" +                                                  // 31 fill alignment
        "</cellXfs>" +
        "<dxfs count=\"9\">" +
        "<dxf><font><color rgb=\"FF9C0006\"/></font><fill><patternFill><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>" +                       // 0 bad
        "<dxf><font><color rgb=\"FF006100\"/></font><fill><patternFill><bgColor rgb=\"FFC6EFCE\"/></patternFill></fill></dxf>" +                       // 1 good
        "<dxf><fill><patternFill><bgColor rgb=\"FFDDEBF7\"/></patternFill></fill></dxf>" +                                                              // 2 zebra
        "<dxf><font><b/><color rgb=\"FF9C5700\"/></font><fill><patternFill><bgColor rgb=\"FFFFEB9C\"/></patternFill></fill>" +
        "<border><bottom style=\"thin\"><color rgb=\"FF9C5700\"/></bottom></border></dxf>" +                                                          // 3 top
        "<dxf><font><b/><i/><u/><color rgb=\"FFC00000\"/></font><numFmt numFmtId=\"167\" formatCode=\"0.0&quot; (!)&quot;\"/></dxf>" +                  // 4 duplicate
        "<dxf><font><b/></font><fill><patternFill><bgColor rgb=\"FFFFD966\"/></patternFill></fill></dxf>" +                                                  // 5 time period
        "<dxf><font><color rgb=\"FF403151\"/></font><fill><patternFill><bgColor rgb=\"FFE4DFEC\"/></patternFill></fill></dxf>" +                      // 6 lavender: the function showcase
        "<dxf><fill><patternFill><bgColor rgb=\"FFF8CBAD\"/></patternFill></fill></dxf>" +                                                              // 7 peach: tables and 3-D references
        "<dxf><fill><patternFill><bgColor rgb=\"FFB4E5A2\"/></patternFill></fill></dxf>" +                                                              // 8 mint: a table reference without a table name
        "</dxfs>" +
        // A table style of the workbook's own for the "Pivot custom" sheet: a yellow header, blue and lavender stripes, a gold total row.
        "<tableStyles count=\"1\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"><tableStyle name=\"Showcase pivot\" table=\"0\" count=\"4\">" +
        "<tableStyleElement type=\"headerRow\" dxfId=\"3\"/><tableStyleElement type=\"firstRowStripe\" dxfId=\"2\"/><tableStyleElement type=\"secondRowStripe\" dxfId=\"6\"/><tableStyleElement type=\"totalRow\" dxfId=\"5\"/></tableStyle></tableStyles>" +
        // The workbook's own palette: entries 0–9 as in Excel, entry 10 (what [Color3] means) purple instead of red.
        "<colors><indexedColors><rgbColor rgb=\"FF000000\"/><rgbColor rgb=\"FFFFFFFF\"/><rgbColor rgb=\"FFFF0000\"/><rgbColor rgb=\"FF00FF00\"/><rgbColor rgb=\"FF0000FF\"/>" +
        "<rgbColor rgb=\"FFFFFF00\"/><rgbColor rgb=\"FFFF00FF\"/><rgbColor rgb=\"FF00FFFF\"/><rgbColor rgb=\"FF000000\"/><rgbColor rgb=\"FFFFFFFF\"/><rgbColor rgb=\"FF7030A0\"/></indexedColors></colors></styleSheet>";

    // ------------------------------------------------------------------ cell helpers

    private static string Str(string reference, string text, int style = Default) =>
        $"<c r=\"{reference}\" t=\"inlineStr\" s=\"{style}\"><is><t xml:space=\"preserve\">{SecurityElement.Escape(text)}</t></is></c>";

    private static string Num(string reference, double value, int style = Default) =>
        $"<c r=\"{reference}\" s=\"{style}\"><v>{value.ToString("R", CultureInfo.InvariantCulture)}</v></c>";

    private static string Row(int row, string cells, double? height = null) =>
        $"<row r=\"{row}\"{(height is { } h ? $" ht=\"{h.ToString(CultureInfo.InvariantCulture)}\" customHeight=\"1\"" : "")}>{cells}</row>";

    private static string ColumnName(int index)
    {
        var name = "";
        for (index++; index > 0; index = (index - 1) / 26) name = (char)('A' + (index - 1) % 26) + name;
        return name;
    }

    /// <summary>An inline string made of formatted runs: each run is its text and the inner XML of its rPr.</summary>
    private static string Rich(string reference, int style, params (string Text, string Properties)[] runs) =>
        $"<c r=\"{reference}\" t=\"inlineStr\" s=\"{style}\"><is>" +
        string.Concat(runs.Select(r => $"<r>{(r.Properties.Length == 0 ? "" : $"<rPr>{r.Properties}</rPr>")}<t xml:space=\"preserve\">{SecurityElement.Escape(r.Text)}</t></r>")) +
        "</is></c>";

    private const string HeaderFooter =
        "<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter></headerFooter>";

    private static string Sheet(string head, string rows, string tail = "") =>
        $"<worksheet xmlns=\"{Main}\" xmlns:r=\"{Rel}\">{head}<sheetData>{rows}</sheetData>{tail}</worksheet>";

    // ------------------------------------------------------------------ sheets

    private static string Formats()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Number and date formats", Title) + Str("B1", "", Title) + Str("C1", "", Title) + Str("D1", "", Title), 30));
        rows.Append(Row(3, Str("A3", "Description", Header) + Str("B3", "Value", Header) + Str("C3", "Format code", Header) + Str("D3", "Notes", Header)));
        var samples = new (string Name, double Value, int Style, string Code, string Note)[]
        {
            ("Thousands separator", 1234.5, Money, "#,##0.00", "1,234.50"),
            ("Percentage", 0.25, Percent, "0%", "25%"),
            ("Percentage, two decimals", 0.2567, Percent2, "0.00%", "25.67%"),
            ("ISO date", 45000, IsoDate, "yyyy-mm-dd", "2023-03-15"),
            ("Long date", 45000, LongDate, "dddd, mmmm d, yyyy", "weekday and month names"),
            ("Time of day", 0.5, Time, "h:mm AM/PM", "noon"),
            ("Negative in red", -1234567, Accounting, "#,##0;[Red](#,##0)", "parentheses"),
            ("Scientific", 12345.678, Scientific, "0.00E+00", "1.23E+04"),
            ("General", 0.1 + 0.2, Default, "General", "no float noise"),
            ("Fraction", 2.5, Fraction1, "# ?/?", "2 1/2"),
            ("Fraction, two digits", 3.14159, Fraction2, "# ??/??", "3 14/99"),
            ("Eighths", 2.375, Eighths, "# ?/8", "2 3/8"),
            ("Indexed color", 1.5, IndexedRed, "[Color3]0.0", "purple, not red"),
        };
        int r = 4;
        foreach (var (name, value, style, code, note) in samples)
        {
            rows.Append(Row(r, Str($"A{r}", name, Boxed) + Num($"B{r}", value, style) + Str($"C{r}", code, Boxed) + Str($"D{r}", note, Boxed)));
            r++;
        }
        rows.Append(Row(r + 1, Str($"A{r + 1}", "Text, boolean and error values are not number-formatted:", Bold)));
        rows.Append(Row(r + 2, Str($"A{r + 2}", "Text", Default) + $"<c r=\"B{r + 2}\" t=\"b\"><v>1</v></c>" + $"<c r=\"C{r + 2}\" t=\"e\"><v>#N/A</v></c>"));

        var cols = "<cols><col min=\"1\" max=\"1\" width=\"24\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"28\" customWidth=\"1\"/>" +
                   "<col min=\"3\" max=\"3\" width=\"20\" customWidth=\"1\"/><col min=\"4\" max=\"4\" width=\"14\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(),
            "<mergeCells count=\"1\"><mergeCell ref=\"A1:D1\"/></mergeCells><printOptions horizontalCentered=\"1\"/>" +
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    private static string LongTable()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Long table report", Title) + Str("B1", "", Title) + Str("C1", "", Title), 26));
        rows.Append(Row(2, Str("A2", "Item", Header) + Str("B2", "Name", Header) + Str("C2", "Amount", Header)));
        for (int i = 1; i <= 200; i++)
        {
            int r = i + 2;
            rows.Append(Row(r, Num($"A{r}", i, Center) + Str($"B{r}", $"Row {i} of the long table") + Num($"C{r}", i * 37.5 % 1000, Money)));
        }
        var cols = "<cols><col min=\"1\" max=\"1\" width=\"10\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"36\" customWidth=\"1\"/><col min=\"3\" max=\"3\" width=\"16\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(),
            "<mergeCells count=\"1\"><mergeCell ref=\"A1:C1\"/></mergeCells>" +
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter +
            "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"60\" max=\"16383\" man=\"1\"/></rowBreaks>");
    }

    private static string Wide()
    {
        var header = new StringBuilder();
        var body = new StringBuilder();
        for (int c = 0; c < 30; c++)
            header.Append(Str($"{ColumnName(c)}1", $"Col {c + 1:00}", Header));
        var rows = new StringBuilder(Row(1, header.ToString()));
        for (int r = 2; r <= 12; r++)
        {
            body.Clear();
            for (int c = 0; c < 30; c++) body.Append(Num($"{ColumnName(c)}{r}", (r - 1) * (c + 1), Default));
            rows.Append(Row(r, body.ToString()));
        }
        return Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"30\" width=\"10\" customWidth=\"1\"/></cols>", rows.ToString(),
            "<pageMargins left=\"0.5\" right=\"0.5\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" +
            "<pageSetup paperSize=\"9\" orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter);
    }

    private static string TextSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Wrapping, overflow, alignment and indentation", Bold)));
        rows.Append(Row(3, Str("A3", "Wrapped text in a narrow column grows the row to fit all of its words.", Wrap) + Str("C3", "Overflow: spills right until a value stops it.") + Str("H3", "stop")));
        rows.Append(Row(5, Str("A5", "Left", Boxed) + Str("B5", "Center", Center) + Str("C5", "Right", Right)));
        rows.Append(Row(6, Str("A6", "Indent 0") ));
        rows.Append(Row(7, Str("A7", "Indent 2", Indent)));
        rows.Append(Row(8, Str("A8", "Too wide for its column:") + Num("B8", 123456789012345, Default)));
        rows.Append(Row(9, Str("A9", "Numbers never spill into neighbors; they show hashes instead.", Default)));
        var cols = "<cols><col min=\"1\" max=\"1\" width=\"24\" customWidth=\"1\"/><col min=\"2\" max=\"7\" width=\"10\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    private static string RichTextSheet()
    {
        const string Base = "<sz val=\"11\"/><rFont val=\"Calibri\"/>";
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Rich text runs inside one cell", Bold)));
        rows.Append(Row(3, Rich("A3", Default, ("Bold", "<b/>" + Base), (", ", Base), ("italic", "<i/>" + Base), (", ", Base), ("underlined", "<u/>" + Base), (", ", Base),
            ("struck", "<strike/>" + Base), (" and ", Base), ("red", "<color rgb=\"FFC00000\"/>" + Base))));
        rows.Append(Row(4, Rich("A4", Default, ("Sizes: ", Base), ("small ", "<sz val=\"8\"/><rFont val=\"Calibri\"/>"), ("LARGE", "<b/><sz val=\"20\"/><color rgb=\"FF1F3864\"/><rFont val=\"Calibri\"/>"))));
        rows.Append(Row(5, Rich("A5", Default, ("H", Base), ("2", "<vertAlign val=\"subscript\"/>" + Base), ("O and E = mc", Base), ("2", "<vertAlign val=\"superscript\"/>" + Base))));
        rows.Append(Row(7, Rich("A7", Wrap, ("Wrapped rich text: ", "<b/>" + Base), ("the runs flow across lines and the row grows to fit them, including a ", Base),
            ("larger", "<sz val=\"16\"/><color rgb=\"FF548235\"/><rFont val=\"Calibri\"/>"), (" word.", Base))));
        rows.Append(Row(9, Rich("A9", Default, ("Overflow: ", "<b/>" + Base), ("rich text spills into empty neighbors until a value stops it.", Base)) + Str("G9", "stop")));
        var cols = "<cols><col min=\"1\" max=\"1\" width=\"30\" customWidth=\"1\"/><col min=\"2\" max=\"7\" width=\"10\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    private static string RotationSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Rotated and stacked text", Bold)));
        rows.Append(Row(3, Str("A3", "Rotate 45", Rot45) + Str("B3", "Rotate 90", Rot90) + Str("C3", "Reading down", RotDown) + Str("D3", "Stacked", Stacked) +
            Str("E3", "A rotated cell that wraps its text onto several lines", RotWrap)));
        rows.Append(Row(4, Str("A4", "Rows grow to fit rotated headings; the text is clipped at the cell edges.")));
        var cols = "<cols><col min=\"1\" max=\"4\" width=\"12\" customWidth=\"1\"/><col min=\"5\" max=\"5\" width=\"22\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    // ------------------------------------------------------------------ shapes

    private static string Fill(string hex) => $"<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>";

    private static string Line(string hex, double width = 1, string dash = "") =>
        $"<a:ln w=\"{(long)(width * Emu)}\">{Fill(hex)}{(dash.Length == 0 ? "" : $"<a:prstDash val=\"{dash}\"/>")}</a:ln>";

    private static string Text(string text, string color = "000000", int size = 1100, bool bold = false, string align = "ctr", string anchor = "ctr") =>
        $"<xdr:txBody><a:bodyPr anchor=\"{anchor}\"/><a:p><a:pPr algn=\"{align}\"/><a:r><a:rPr lang=\"en-US\" sz=\"{size}\" b=\"{(bold ? 1 : 0)}\">{Fill(color)}</a:rPr><a:t>{SecurityElement.Escape(text)}</a:t></a:r></a:p></xdr:txBody>";

    private static string Shape(string preset, string name, string fill, string extra = "", string text = "", int rotation = 0, string flip = "") =>
        $"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"{name}\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm rot=\"{rotation}\"{flip}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{(fill.Length == 0 ? "<a:noFill/>" : Fill(fill))}{extra}</xdr:spPr>{text}</xdr:sp>";

    private static void WriteShapes(ZipArchive zip, int sheetNumber)
    {
        // The default look Excel gives an inserted shape: no fill or line in spPr, the theme's blue and a darker outline from the style.
        const string ThemeStyle = "<xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>" +
                                  "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></xdr:style>";
        string themed = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Themed\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>" + ThemeStyle + Text("Theme style", "FFFFFF") + "</xdr:sp>";
        string connector(string flip, string color) =>
            $"<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"4\" name=\"Line\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm{flip}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></a:xfrm><a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom>{Line(color, 2.25)}</xdr:spPr></xdr:cxnSp>";
        string child(string preset, string fill, int x, int y, int w, int h) =>
            $"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"5\" name=\"Part\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm><a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{Fill(fill)}</xdr:spPr></xdr:sp>";
        var group = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"6\" name=\"Group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2000\" cy=\"1000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"2000\" cy=\"1000\"/></a:xfrm></xdr:grpSpPr>" +
                    child("ellipse", "ED7D31", 0, 0, 1000, 1000) + child("rect", "70AD47", 1000, 250, 1000, 500) + "</xdr:grpSp>";
        var anchors =
            OneCell(1, 3, 0, 0, 130, 50, themed) +
            OneCell(4, 3, 0, 0, 130, 50, Shape("roundRect", "Rounded", "FFF2CC", Line("BF9000", 1.5), Text("Rounded rectangle", "7F6000", bold: true))) +
            OneCell(7, 3, 0, 0, 100, 50, Shape("ellipse", "Oval", "DDEBF7", Line("2F5597", 1, "dash"), Text("Dashed oval", "1F3864"))) +
            OneCell(1, 8, 0, 0, 90, 60, Shape("triangle", "Triangle", "C00000", "", Text("Up", "FFFFFF"))) +
            OneCell(3, 8, 0, 0, 90, 60, Shape("diamond", "Diamond", "7030A0", "", Text("Diamond", "FFFFFF"))) +
            OneCell(5, 8, 0, 0, 110, 50, Shape("rightArrow", "Arrow", "00B050", Line("006100"), Text("Right", "FFFFFF"))) +
            OneCell(7, 8, 0, 0, 90, 60, Shape("hexagon", "Hexagon", "FFC000", Line("7F6000", 2), Text("Hexagon"))) +
            OneCell(1, 14, 0, 0, 120, 45, Shape("rect", "Rotated", "9DC3E6", Line("2E75B6", 1), Text("Rotated 20°", "1F3864"), rotation: 20 * 60000)) +
            OneCell(4, 14, 0, 0, 110, 45, Shape("rect", "Text box", "", "", Text("A text box has no fill or outline, only text.", "404040", 1000, align: "l", anchor: "t"))) +
            OneCell(7, 14, 0, 0, 100, 50, group) +
            OneCell(1, 20, 0, 0, 120, 40, connector("", "C00000")) +
            OneCell(4, 20, 0, 0, 120, 40, connector(" flipV=\"1\"", "2F5597")) +
            OneCell(7, 20, 0, 0, 100, 0, connector("", "7F7F7F"));
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
    }

    private static string ShapesSheet()
    {
        var rows = Row(1, Str("A1", "Shapes, text boxes and lines", Bold)) + Row(3, Str("A3", "Shapes:")) + Row(8, Str("A8", "Geometry:")) + Row(12, Str("A12", "Rotation, text box, group:")) +
                   Row(20, Str("A20", "Lines:"));
        return Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"10\" width=\"13\" customWidth=\"1\"/></cols>", rows,
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");
    }

    private static string Gradient(string from, string to, string direction) =>
        $"<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"{from}\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"{to}\"/></a:gs></a:gsLst>{direction}</a:gradFill>";

    private static string Custom(string name, string path, string fill, string extra, string text = "") =>
        $"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"{name}\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
        $"<a:custGeom><a:avLst/><a:pathLst><a:path w=\"100\" h=\"100\">{path}</a:path></a:pathLst></a:custGeom>{fill}{extra}</xdr:spPr>{text}</xdr:sp>";

    private static string Connector(string preset, string color, string ends, string adjust = "", string flip = "") =>
        $"<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"4\" name=\"Connector\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm{flip}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"{preset}\"><a:avLst>{adjust}</a:avLst></a:prstGeom><a:ln w=\"{(long)(2.25 * Emu)}\">{Fill(color)}{ends}</a:ln></xdr:spPr></xdr:cxnSp>";

    private static void WriteShapePaths(ZipArchive zip, int sheetNumber)
    {
        string Pt(int x, int y) => $"<a:pt x=\"{x}\" y=\"{y}\"/>";
        var star = $"<a:moveTo>{Pt(50, 0)}</a:moveTo><a:lnTo>{Pt(62, 38)}</a:lnTo><a:lnTo>{Pt(100, 38)}</a:lnTo><a:lnTo>{Pt(69, 61)}</a:lnTo><a:lnTo>{Pt(81, 100)}</a:lnTo><a:lnTo>{Pt(50, 76)}</a:lnTo>" +
                   $"<a:lnTo>{Pt(19, 100)}</a:lnTo><a:lnTo>{Pt(31, 61)}</a:lnTo><a:lnTo>{Pt(0, 38)}</a:lnTo><a:lnTo>{Pt(38, 38)}</a:lnTo><a:close/>";
        var heart = $"<a:moveTo>{Pt(50, 30)}</a:moveTo><a:cubicBezTo>{Pt(50, 0)}{Pt(0, 0)}{Pt(0, 35)}</a:cubicBezTo><a:cubicBezTo>{Pt(0, 65)}{Pt(35, 80)}{Pt(50, 100)}</a:cubicBezTo>" +
                    $"<a:cubicBezTo>{Pt(65, 80)}{Pt(100, 65)}{Pt(100, 35)}</a:cubicBezTo><a:cubicBezTo>{Pt(100, 0)}{Pt(50, 0)}{Pt(50, 30)}</a:cubicBezTo><a:close/>";
        var wave = $"<a:moveTo>{Pt(0, 50)}</a:moveTo><a:cubicBezTo>{Pt(20, 0)}{Pt(30, 0)}{Pt(50, 50)}</a:cubicBezTo><a:cubicBezTo>{Pt(70, 100)}{Pt(80, 100)}{Pt(100, 50)}</a:cubicBezTo>";
        string End(string head, string tail) => (head.Length == 0 ? "" : $"<a:headEnd type=\"{head}\" w=\"med\" len=\"med\"/>") + (tail.Length == 0 ? "" : $"<a:tailEnd type=\"{tail}\" w=\"med\" len=\"med\"/>");
        var anchors =
            OneCell(1, 3, 0, 0, 110, 55, Shape("rect", "Linear", "", Line("1F3864"), Text("Linear", "FFFFFF", bold: true)).Replace("<a:noFill/>", Gradient("1F3864", "8FAADC", "<a:lin ang=\"0\"/>"))) +
            OneCell(3, 3, 0, 0, 90, 55, Shape("ellipse", "Radial", "", Line("7F6000"), Text("Radial", "000000", bold: true)).Replace("<a:noFill/>", Gradient("FFF2CC", "ED7D31", "<a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path>"))) +
            OneCell(5, 3, 0, 0, 110, 55, Shape("roundRect", "Rounded", "", Line("375623"), Text("Rounded", "FFFFFF", bold: true)).Replace("<a:noFill/>", Gradient("A9D18E", "375623", "<a:lin ang=\"5400000\"/>"))) +
            OneCell(7, 3, 0, 0, 90, 55, Shape("diamond", "Diagonal", "", "", Text("45°", "FFFFFF", bold: true)).Replace("<a:noFill/>", Gradient("C00000", "FFC000", "<a:lin ang=\"2700000\"/>"))) +
            OneCell(1, 9, 0, 0, 110, 0, Connector("straightConnector1", "C00000", End("", "triangle"))) +
            OneCell(3, 9, 0, 0, 110, 0, Connector("straightConnector1", "2F5597", End("oval", "stealth"))) +
            OneCell(5, 9, 0, 0, 110, 0, Connector("straightConnector1", "375623", End("diamond", "arrow"))) +
            OneCell(7, 9, 0, 0, 90, 30, Connector("straightConnector1", "7030A0", End("triangle", "triangle"), flip: " flipV=\"1\"")) +
            OneCell(1, 13, 0, 0, 110, 50, Connector("bentConnector3", "C00000", End("oval", "triangle"))) +
            OneCell(3, 13, 0, 0, 110, 50, Connector("bentConnector3", "2F5597", End("", "triangle"), "<a:gd name=\"adj1\" fmla=\"val 20000\"/>")) +
            OneCell(5, 13, 0, 0, 110, 50, Connector("bentConnector2", "375623", End("", "stealth"))) +
            OneCell(7, 13, 0, 0, 90, 50, Connector("curvedConnector3", "7030A0", End("oval", "triangle"))) +
            OneCell(1, 19, 0, 0, 90, 80, Custom("Star", star, Fill("FFC000"), Line("7F6000", 1.5))) +
            OneCell(3, 19, 0, 0, 90, 80, Custom("Heart", heart, Gradient("FF7C80", "C00000", "<a:lin ang=\"5400000\"/>"), Line("7F0000", 1))) +
            OneCell(5, 19, 0, 0, 110, 60, Custom("Wave", wave, "", Line("2F5597", 2.5).Replace("</a:ln>", End("", "triangle") + "</a:ln>"))) +
            OneCell(9, 19, 0, 0, 90, 80, Custom("Frame with a hole", $"<a:moveTo>{Pt(0, 0)}</a:moveTo><a:lnTo>{Pt(100, 0)}</a:lnTo><a:lnTo>{Pt(100, 100)}</a:lnTo><a:lnTo>{Pt(0, 100)}</a:lnTo><a:close/>" +
                                                      $"<a:moveTo>{Pt(25, 25)}</a:moveTo><a:lnTo>{Pt(25, 75)}</a:lnTo><a:lnTo>{Pt(75, 75)}</a:lnTo><a:lnTo>{Pt(75, 25)}</a:lnTo><a:close/>", Fill("7030A0"), Line("3B1A55", 1.5))) +
            OneCell(7, 19, 0, 0, 100, 70, Custom("Star with text", star, Fill("DDEBF7"), Line("2F5597", 1.5), Text("Custom", "1F3864", bold: true))) +
            OneCell(1, 26, 0, 0, 90, 60, Shape("pentagon", "Pentagon", "2F5597", Line("1F3864"), Text("Pentagon", "FFFFFF", 900, bold: true))) +
            OneCell(3, 26, 0, 0, 90, 60, Shape("octagon", "Octagon", "C00000", Line("7F0000"), Text("Octagon", "FFFFFF", 900, bold: true))) +
            OneCell(5, 26, 0, 0, 90, 60, Shape("plus", "Plus", "70AD47", Line("375623"), Text("Plus", "FFFFFF", 900, bold: true))) +
            OneCell(7, 26, 0, 0, 110, 60, Shape("chevron", "Chevron", "ED7D31", Line("843C0C"), Text("Chevron", "FFFFFF", 900, bold: true))) +
            OneCell(1, 32, 0, 0, 110, 60, Shape("homePlate", "Pentagon arrow", "7030A0", Line("3B1A55"), Text("Pentagon arrow", "FFFFFF", 900, bold: true))) +
            OneCell(3, 32, 0, 0, 90, 80, Shape("star5", "Star", "FFC000", Line("7F6000"), Text("Star", "000000", 900, bold: true))) +
            OneCell(5, 32, 0, 0, 100, 60, Shape("roundRect", "Round corners 40%", "5B9BD5", Line("1F4E79"), Text("Round 40%", "FFFFFF", 900, bold: true)).Replace("<a:avLst/>", "<a:avLst><a:gd name=\"adj\" fmla=\"val 40000\"/></a:avLst>")) +
            OneCell(7, 32, 0, 0, 120, 60, Shape("rightArrow", "Slim arrow", "548235", Line("375623"), Text("Slim arrow", "FFFFFF", 900, bold: true)).Replace("<a:avLst/>", "<a:avLst><a:gd name=\"adj1\" fmla=\"val 30000\"/><a:gd name=\"adj2\" fmla=\"val 80000\"/></a:avLst>"));
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
    }

    private static string ShapePathsSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"10\" width=\"13\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "Shape paths: gradients, arrowheads, bent connectors, custom geometry", Bold)) + Row(3, Str("A3", "Gradient fills:")) + Row(9, Str("A9", "Arrowheads:")) + Row(13, Str("A13", "Connectors:")) + Row(19, Str("A19", "Custom geometry:")) + Row(26, Str("A26", "More presets and adjust values:")),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    private static readonly string[] PresetNames =
    [
        "star4", "star6", "star8", "star10", "star24", "heptagon", "decagon", "dodecagon",
        "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "borderCallout1", "cloud",
        "flowChartDocument", "flowChartPredefinedProcess", "flowChartInternalStorage", "flowChartManualInput", "flowChartManualOperation", "flowChartDelay",
        "flowChartMerge", "flowChartDisplay", "flowChartOffpageConnector", "flowChartOr", "flowChartSummingJunction", "flowChartMagneticDisk",
        "leftBracket", "rightBracket", "leftBrace", "rightBrace", "can", "cube", "bevel", "foldedCorner",
        "heart", "lightningBolt", "sun", "moon", "smileyFace", "teardrop", "leftRightArrow", "upDownArrow",
        "donut", "frame", "noSmoking", "mathPlus", "mathMinus", "mathEqual", "snip1Rect", "snip2SameRect", "round1Rect", "round2SameRect", "pie", "chord", "arc", "blockArc",
    ];

    private static string ShapePresetsSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"9\" width=\"13\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "Shape presets: stars, callouts, flow charts, brackets, solids, symbols", Bold)),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    private static void WriteShapePresets(ZipArchive zip, int sheetNumber)
    {
        var anchors = new StringBuilder();
        for (int i = 0; i < PresetNames.Length; i++)
        {
            // Eight to a row, each in a box with the name below it; the callouts leave room for their tips.
            int column = i % 8, row = 3 + i / 8 * 6;
            var name = PresetNames[i];
            var fill = new[] { "2F5597", "C55A11", "548235", "7030A0", "BF9000", "C00000" }[i / 8 % 6];
            anchors.Append(OneCell(column, row, 4, 4, 70, 50, Shape(name, name, fill, Line("203864", 1))));
            anchors.Append(OneCell(column, row + 4, 0, 6, 92, 14, Shape("rect", "Label " + name, "", "", Text(name.Replace("flowChart", ""), "000000", 600))));
        }
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
    }

    private static string Pattern(string preset, string foreground, string background) =>
        $"<a:pattFill prst=\"{preset}\"><a:fgClr><a:srgbClr val=\"{foreground}\"/></a:fgClr><a:bgClr><a:srgbClr val=\"{background}\"/></a:bgClr></a:pattFill>";

    private static void WriteShapeFills(ZipArchive zip, int sheetNumber)
    {
        string Style(int fill, int effect, string color = "accent1") =>
            $"<xdr:style><a:lnRef idx=\"0\"><a:schemeClr val=\"{color}\"/></a:lnRef><a:fillRef idx=\"{fill}\"><a:schemeClr val=\"{color}\"/></a:fillRef><a:effectRef idx=\"{effect}\"><a:schemeClr val=\"{color}\"/></a:effectRef>" +
            "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></xdr:style>";
        string Styled(string name, string preset, string style, string text) =>
            $"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"{name}\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom></xdr:spPr>{style}{text}</xdr:sp>";
        string Hatched(string preset, string name, string shape = "rect") =>
            Shape(shape, name, "", Line("C55A11", 1)).Replace("<a:noFill/>", Pattern(preset, "C55A11", "FFF2CC")).Replace("</xdr:sp>", Text(preset, "000000", 900) + "</xdr:sp>");
        string Shadowed(string name, string preset, string fill, string effects) =>
            Shape(preset, name, fill, Line("404040", 1) + effects, Text(name, "FFFFFF", bold: true));
        const string Drop = "<a:effectLst><a:outerShdw blurRad=\"50800\" dist=\"50800\" dir=\"2700000\" algn=\"tl\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"40000\"/></a:srgbClr></a:outerShdw></a:effectLst>";
        // A chevron whose point is a guide formula: the notch depth is a fraction (adj) of the shorter side.
        var chevron = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"Chevron\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1270000\" cy=\"635000\"/></a:xfrm>" +
                      "<a:custGeom><a:avLst><a:gd name=\"adj\" fmla=\"val 50000\"/></a:avLst><a:gdLst><a:gd name=\"maxAdj\" fmla=\"*/ 100000 w ss\"/><a:gd name=\"a\" fmla=\"pin 0 adj maxAdj\"/>" +
                      "<a:gd name=\"x1\" fmla=\"*/ ss a 100000\"/><a:gd name=\"x2\" fmla=\"+- r 0 x1\"/><a:gd name=\"x3\" fmla=\"*/ x2 1 2\"/></a:gdLst>" +
                      "<a:pathLst><a:path><a:moveTo><a:pt x=\"l\" y=\"t\"/></a:moveTo><a:lnTo><a:pt x=\"x2\" y=\"t\"/></a:lnTo><a:lnTo><a:pt x=\"r\" y=\"vc\"/></a:lnTo><a:lnTo><a:pt x=\"x2\" y=\"b\"/></a:lnTo><a:lnTo><a:pt x=\"l\" y=\"b\"/></a:lnTo><a:lnTo><a:pt x=\"x1\" y=\"vc\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>" +
                      Fill("70AD47") + Line("375623", 1.5) + "</xdr:spPr>" + Text("Guides", "FFFFFF", bold: true) + "</xdr:sp>";
        string[] presets = ["ltUpDiag", "dkHorz", "narVert", "wdDnDiag", "cross", "diagCross", "pct25", "smGrid"];
        var anchors = new StringBuilder();
        for (int i = 0; i < presets.Length; i++)
            anchors.Append(OneCell(1 + i % 4 * 2, 3 + i / 4 * 5, 0, 0, 100, 55, Hatched(presets[i], presets[i], i == 3 ? "ellipse" : i == 5 ? "hexagon" : "rect")));
        anchors.Append(OneCell(1, 14, 0, 0, 110, 55, Styled("Gradient 2", "rect", Style(2, 0), Text("Style gradient 2", "FFFFFF", bold: true))));
        anchors.Append(OneCell(3, 14, 0, 0, 110, 55, Styled("Gradient 3", "roundRect", Style(3, 0, "accent2"), Text("Style gradient 3", "FFFFFF", bold: true))));
        anchors.Append(OneCell(5, 14, 0, 0, 110, 55, Styled("Themed shadow", "rect", Style(1, 3, "accent6"), Text("Theme shadow", "FFFFFF", bold: true))));
        anchors.Append(OneCell(7, 14, 0, 0, 100, 55, Shadowed("Drop shadow", "ellipse", "7030A0", Drop)));
        anchors.Append(OneCell(1, 20, 0, 0, 100, 50, chevron));
        // Soft shadows, picture fills and rectangular path gradients.
        const string Blurred = "<a:effectLst><a:outerShdw blurRad=\"190500\" dist=\"76200\" dir=\"2700000\" algn=\"tl\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"60000\"/></a:srgbClr></a:outerShdw></a:effectLst>";
        string Stops(string a, string b) => $"<a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"{a}\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"{b}\"/></a:gs></a:gsLst>";
        string PathGradient(string path, string a, string b) =>
            $"<a:gradFill>{Stops(a, b)}<a:path path=\"{path}\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill>";
        string Picture(string preset, string name) =>
            Shape(preset, name, "", Line("404040", 1), Text(name, "000000", 900, bold: true, anchor: "b")).Replace("<a:noFill/>", "<a:blipFill><a:blip r:embed=\"rId1\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>");
        anchors.Append(OneCell(3, 20, 0, 0, 100, 50, Shadowed("Blurred shadow", "roundRect", "ED7D31", Blurred)));
        anchors.Append(OneCell(5, 20, 0, 0, 100, 50, Picture("rect", "Picture fill")));
        anchors.Append(OneCell(7, 20, 0, 0, 100, 50, Picture("ellipse", "Cut to an ellipse")));
        anchors.Append(OneCell(1, 25, 0, 0, 100, 50, Shape("rect", "Rectangular", "").Replace("<a:noFill/>", PathGradient("rect", "FFFFFF", "C00000")).Replace("</xdr:sp>", Text("Rectangular", "000000", 900, bold: true) + "</xdr:sp>")));
        anchors.Append(OneCell(3, 25, 0, 0, 100, 50, Shape("diamond", "Rect in a diamond", "").Replace("<a:noFill/>", PathGradient("rect", "FFC000", "7030A0")).Replace("</xdr:sp>", Text("Diamond", "FFFFFF", 900, bold: true) + "</xdr:sp>")));
        anchors.Append(OneCell(5, 25, 0, 0, 100, 50, Shape("rect", "Circular", "").Replace("<a:noFill/>", PathGradient("circle", "FFFFFF", "2F5597")).Replace("</xdr:sp>", Text("Circular", "000000", 900, bold: true) + "</xdr:sp>")));
        // Tiled and translucent picture fills, a glow, soft edges and a rotated group.
        string Blip(string inside, string rest) => $"<a:blipFill><a:blip r:embed=\"rId1\">{inside}</a:blip>{rest}</a:blipFill>";
        anchors.Append(OneCell(1, 31, 0, 0, 100, 60, Shape("rect", "Tiled picture", "", Line("404040", 1), Text("Tiled picture", "000000", 900, bold: true, anchor: "b")).Replace("<a:noFill/>", Blip("", "<a:tile tx=\"0\" ty=\"0\" sx=\"30000\" sy=\"30000\" algn=\"tl\" flip=\"none\"/>"))));
        anchors.Append(OneCell(5, 37, 0, 0, 100, 60, Shape("rect", "Mirrored tiles", "", Line("404040", 1), Text("Mirrored tiles", "000000", 900, bold: true, anchor: "b")).Replace("<a:noFill/>", Blip("", "<a:tile tx=\"0\" ty=\"0\" sx=\"30000\" sy=\"30000\" algn=\"tl\" flip=\"xy\"/>"))));
        anchors.Append(OneCell(3, 31, 0, 0, 100, 60, Shape("rect", "Translucent picture", "", Line("404040", 1), Text("Translucent picture", "000000", 900, bold: true, anchor: "b")).Replace("<a:noFill/>", Blip("<a:alphaModFix amt=\"45000\"/>", "<a:stretch><a:fillRect/></a:stretch>"))));
        anchors.Append(OneCell(5, 31, 0, 0, 100, 50, Shadowed("Glow", "ellipse", "5B9BD5", "<a:effectLst><a:glow rad=\"101600\"><a:srgbClr val=\"FFC000\"><a:alpha val=\"45000\"/></a:srgbClr></a:glow></a:effectLst>")));
        // A glow outside a translucent picture fill: it must not show through the fill.
        anchors.Append(OneCell(7, 25, 6, 6, 88, 44, Shape("ellipse", "Outside glow", "", Line("404040", 1), Text("Outside glow", "000000", 900, bold: true, anchor: "b") .Replace("</xdr:txBody>", "</xdr:txBody>"))
            .Replace("<a:noFill/>", Blip("<a:alphaModFix amt=\"35000\"/>", "<a:stretch><a:fillRect/></a:stretch>"))
            .Replace("</xdr:spPr>", "<a:effectLst><a:glow rad=\"127000\"><a:srgbClr val=\"ED7D31\"><a:alpha val=\"60000\"/></a:srgbClr></a:glow></a:effectLst></xdr:spPr>")));
        anchors.Append(OneCell(7, 31, 0, 0, 100, 50, Shadowed("Soft edges", "rect", "C00000", "<a:effectLst><a:softEdge rad=\"127000\"/></a:effectLst>")));
        string Child(string name, string fill, int x, int y, int cx, int cy) =>
            $"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"{name}\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{Fill(fill)}</xdr:spPr>{Text(name, "FFFFFF", 900, bold: true)}</xdr:sp>";
        anchors.Append(OneCell(1, 37, 0, 0, 200, 70,
            "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"9\" name=\"Rotated group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm rot=\"1200000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2540000\" cy=\"889000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"2540000\" cy=\"889000\"/></a:xfrm></xdr:grpSpPr>" +
            Child("Group A", "2F5597", 0, 0, 1000000, 889000) + Child("Group B", "70AD47", 1270000, 0, 1270000, 444500) + "</xdr:grpSp>"));
        // A reflection under a shape, soft edges that follow a notched outline, and curved connectors with four and five segments.
        const string Reflect = "<a:effectLst><a:reflection blurRad=\"0\" stA=\"55000\" endA=\"0\" endPos=\"80000\" dist=\"0\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/></a:effectLst>";
        anchors.Append(OneCell(1, 43, 0, 0, 100, 44, Shape("roundRect", "Reflection", "2F5597", Line("1F3864", 1) + Reflect, Text("Reflection", "FFFFFF", bold: true))));
        string Lt(int x, int y) => $"<a:lnTo><a:pt x=\"{x}\" y=\"{y}\"/></a:lnTo>";
        var notch = "<a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>" + Lt(100, 0) + Lt(100, 40) + Lt(40, 40) + Lt(40, 100) + Lt(0, 100) + "<a:close/>";
        anchors.Append(OneCell(3, 43, 0, 0, 90, 70, Custom("Notched soft edges", notch, Fill("C00000"), "<a:effectLst><a:softEdge rad=\"190500\"/></a:effectLst>", Text("Notched soft edges", "FFFFFF", 800, bold: true))));
        const string Arrow = "<a:tailEnd type=\"triangle\" w=\"med\" len=\"med\"/>";
        anchors.Append(OneCell(5, 43, 0, 0, 90, 60, Connector("curvedConnector4", "7030A0", Arrow, "<a:gd name=\"adj1\" fmla=\"val 30000\"/><a:gd name=\"adj2\" fmla=\"val 70000\"/>")));
        anchors.Append(OneCell(7, 43, 0, 0, 90, 60, Connector("curvedConnector5", "375623", Arrow, "<a:gd name=\"adj1\" fmla=\"val 20000\"/><a:gd name=\"adj2\" fmla=\"val 50000\"/><a:gd name=\"adj3\" fmla=\"val 80000\"/>")));
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/drawings/_rels/drawing{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}image\" Target=\"../media/shapefill.png\"/></Relationships>");
        using (var media = zip.CreateEntry("xl/media/shapefill.png").Open()) media.Write(ShowcaseDocx.GradientPng(120, 80));
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
    }

    private static string ShapeFillsSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"10\" width=\"13\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "Shape fills: patterns, theme gradients, shadows, guide formulas", Bold)) + Row(3, Str("A3", "Pattern fills:")) + Row(14, Str("A14", "Theme style gradients and shadows:")) + Row(20, Str("A20", "Guides, soft shadow, picture fills:")) + Row(25, Str("A25", "Rectangular and circular path gradients:")) + Row(31, Str("A31", "Tiled and translucent picture fills, glow, soft edges:")) + Row(37, Str("A37", "A rotated group:")) + Row(43, Str("A43", "Reflection, notched soft edges, curved connectors 4 and 5:")) + Row(53, Str("A53", "(the print range ends here)")),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    // ------------------------------------------------------------------ charts

    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private static readonly string[] Quarters = ["Q1", "Q2", "Q3", "Q4", "Q5", "Q6"];
    private static readonly double[][] ChartData = [[12, 19, 28, 24, 35, 41], [8, 14, 11, 22, 18, 30], [5, 9, 16, 12, 20, 15]];
    private static readonly string[] SeriesNames = ["North", "South", "East"];

    // The stock chart's close: above the open (series 0) in four categories, below it in two, so white and black bodies both appear.
    private static readonly double[] StockClose = [14, 17, 30, 20, 38, 36];

    private static string ChartSeries(int i, bool cached, string extra = "", double[]? data = null)
    {
        string column = ((char)('B' + i)).ToString();
        string name = cached ? $"<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{SeriesNames[i]}</c:v></c:pt></c:strCache>" : "";
        string cats = cached ? "<c:strCache><c:ptCount val=\"6\"/>" + string.Concat(Quarters.Select((q, k) => $"<c:pt idx=\"{k}\"><c:v>{q}</c:v></c:pt>")) + "</c:strCache>" : "";
        string vals = cached ? "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"6\"/>" + string.Concat((data ?? ChartData[i]).Select((v, k) => $"<c:pt idx=\"{k}\"><c:v>{v}</c:v></c:pt>")) + "</c:numCache>" : "";
        return $"<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/><c:tx><c:strRef><c:f>Charts!${column}$3</c:f>{name}</c:strRef></c:tx>{extra}" +
               $"<c:cat><c:strRef><c:f>Charts!$A$4:$A$9</c:f>{cats}</c:strRef></c:cat><c:val><c:numRef><c:f>Charts!${column}$4:${column}$9</c:f>{vals}</c:numRef></c:val></c:ser>";
    }

    private static string Rich(string tag, string text, int size = 1400) =>
        $"<c:{tag}><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"{size}\"/><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:{tag}>";

    private static string ChartSpace(string title, string plot, string legend = "r") =>
        $"<c:chartSpace xmlns:c=\"{ChartNs}\" xmlns:a=\"{DmlNs}\"><c:chart>{(title.Length == 0 ? "<c:autoTitleDeleted val=\"1\"/>" : Rich("title", title))}<c:plotArea><c:layout/>{plot}</c:plotArea>" +
        $"{(legend.Length == 0 ? "" : $"<c:legend><c:legendPos val=\"{legend}\"/><c:overlay val=\"0\"/></c:legend>")}</c:chart></c:chartSpace>";

    private static string Axes(string catTitle = "", string valTitle = "") =>
        $"<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>{(catTitle.Length == 0 ? "" : Rich("title", catTitle, 1000))}<c:crossAx val=\"2\"/></c:catAx>" +
        $"<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/>{(valTitle.Length == 0 ? "" : Rich("title", valTitle, 1000))}<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:crossAx val=\"1\"/></c:valAx>";

    private static string Bars(string dir, string grouping, int count, bool cached, string extra = "", string axes = "") =>
        $"<c:barChart><c:barDir val=\"{dir}\"/><c:grouping val=\"{grouping}\"/><c:varyColors val=\"0\"/>{string.Concat(Enumerable.Range(0, count).Select(i => ChartSeries(i, cached)))}{extra}" +
        $"<c:gapWidth val=\"150\"/>{(grouping == "clustered" ? "" : "<c:overlap val=\"100\"/>")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>{(axes.Length == 0 ? Axes() : axes)}";

    private static void WriteCharts(ZipArchive zip, int sheetNumber)
    {
        const string Labels = "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbls>";
        const string NoMarker = "<c:marker><c:symbol val=\"none\"/></c:marker>";
        var line = "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + ChartSeries(0, true) + ChartSeries(1, true, NoMarker) + "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" + Axes();
        var pie = "<c:pieChart><c:varyColors val=\"1\"/>" + ChartSeries(0, true).Replace("</c:ser>", "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/></c:dLbls></c:ser>") + "</c:pieChart>";
        const string Unjoined = "<c:spPr><a:ln w=\"19050\"><a:noFill/></a:ln></c:spPr>";
        var stock = "<c:stockChart>" + ChartSeries(0, true, Unjoined + NoMarker) + ChartSeries(1, true, Unjoined + NoMarker) + ChartSeries(2, true, Unjoined + "<c:marker><c:symbol val=\"dash\"/><c:size val=\"7\"/></c:marker>", StockClose) +
                    "<c:hiLowLines/><c:upDownBars><c:gapWidth val=\"150\"/></c:upDownBars><c:axId val=\"1\"/><c:axId val=\"2\"/></c:stockChart>" + Axes();
        var surface = "<c:surfaceChart/>";
        string[] charts =
        [
            ChartSpace("Clustered columns", Bars("col", "clustered", 3, true, axes: Axes("Quarter", "Units")), "b"),
            ChartSpace("Stacked columns", Bars("col", "stacked", 3, false)),
            ChartSpace("Bars with labels", Bars("bar", "clustered", 1, true, Labels), ""),
            ChartSpace("Lines", line, "b"),
            ChartSpace("Pie shares", pie),
            ChartSpace("100 % stacked", Bars("col", "percentStacked", 3, true)),
            ChartSpace("Stock", stock, "b"), ChartSpace("Surface", surface),
        ];
        (int C0, int R0, int C1, int R1)[] places = [(0, 11, 3, 24), (3, 11, 6, 24), (0, 25, 3, 38), (3, 25, 6, 38), (0, 39, 3, 52), (3, 39, 6, 52), (0, 53, 3, 66), (3, 53, 6, 66)];
        WriteChartParts(zip, sheetNumber, charts, places, "chart");
    }

    private static void WriteChartParts(ZipArchive zip, int sheetNumber, string[] charts, (int C0, int R0, int C1, int R1)[] places, string prefix)
    {
        var anchors = new StringBuilder();
        var rels = new StringBuilder();
        for (int i = 0; i < charts.Length; i++)
        {
            Entry(zip, $"xl/charts/{prefix}{sheetNumber}_{i + 1}.xml", charts[i]);
            rels.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"{RelType}chart\" Target=\"../charts/{prefix}{sheetNumber}_{i + 1}.xml\"/>");
            var (c0, r0, c1, r1) = places[i];
            anchors.Append(TwoCell(c0, r0, c1, r1,
                "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + (i + 2) + "\" name=\"Chart " + (i + 1) + "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>" +
                $"<a:graphic><a:graphicData uri=\"{ChartNs}\"><c:chart xmlns:c=\"{ChartNs}\" r:id=\"rId{i + 1}\"/></a:graphicData></a:graphic></xdr:graphicFrame>"));
        }
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/drawings/_rels/drawing{sheetNumber}.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
    }

    private static string ChartsSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Charts: columns, bars, lines and pies", Bold)));
        rows.Append(Row(3, Str("A3", "Quarter", Header) + Str("B3", SeriesNames[0], Header) + Str("C3", SeriesNames[1], Header) + Str("D3", SeriesNames[2], Header)));
        for (int k = 0; k < 6; k++)
            rows.Append(Row(4 + k, Str($"A{4 + k}", Quarters[k]) + Num($"B{4 + k}", ChartData[0][k]) + Num($"C{4 + k}", ChartData[1][k]) + Num($"D{4 + k}", ChartData[2][k])));
        rows.Append(Row(11, Str("A11", "The surface chart at the bottom is not supported and is reported.")));
        return Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"6\" width=\"16\" customWidth=\"1\"/></cols>", rows.ToString(),
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"1\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");
    }

    private static string XySeries(int i, string xs, string ys, string extra = "", string sizes = "")
    {
        string Cache(double[] v) => "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"" + v.Length + "\"/>" + string.Concat(v.Select((d, k) => $"<c:pt idx=\"{k}\"><c:v>{d}</c:v></c:pt>")) + "</c:numCache>";
        double[] x = ChartData[0], y = ChartData[1], z = ChartData[2];
        return $"<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/><c:tx><c:v>{SeriesNames[i]}</c:v></c:tx>{extra}<c:xVal><c:numRef><c:f>{xs}</c:f>{Cache(x)}</c:numRef></c:xVal>" +
               $"<c:yVal><c:numRef><c:f>{ys}</c:f>{Cache(y)}</c:numRef></c:yVal>{(sizes.Length == 0 ? "" : $"<c:bubbleSize><c:numRef><c:f>{sizes}</c:f>{Cache(z)}</c:numRef></c:bubbleSize>")}</c:ser>";
    }

    private const string XyAxes =
        "<c:valAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:majorGridlines/><c:crossAx val=\"2\"/></c:valAx>" +
        "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:crossAx val=\"1\"/></c:valAx>";

    private static void WriteMoreCharts(ZipArchive zip, int sheetNumber)
    {
        const string MarkersOnly = "<c:spPr><a:ln w=\"19050\"><a:noFill/></a:ln></c:spPr>";
        const string Percent = "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/></c:dLbls>";
        var area = "<c:areaChart><c:grouping val=\"stacked\"/><c:varyColors val=\"0\"/>" + ChartSeries(0, true) + ChartSeries(1, true) + ChartSeries(2, true) + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>" +
                   Axes().Replace("<c:crossAx val=\"1\"/></c:valAx>", "<c:crossAx val=\"1\"/><c:crossBetween val=\"midCat\"/></c:valAx>");
        var doughnut = "<c:doughnutChart><c:varyColors val=\"1\"/>" + ChartSeries(0, true).Replace("</c:ser>", Percent + "</c:ser>") + "<c:firstSliceAng val=\"0\"/><c:holeSize val=\"55\"/></c:doughnutChart>";
        var rings = "<c:doughnutChart><c:varyColors val=\"1\"/>" + ChartSeries(0, true) + ChartSeries(1, true) + "<c:firstSliceAng val=\"0\"/><c:holeSize val=\"40\"/></c:doughnutChart>";
        var scatter = "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>" + XySeries(0, "Charts!$B$4:$B$9", "Charts!$C$4:$C$9") + XySeries(1, "Charts!$B$4:$B$9", "Charts!$D$4:$D$9", MarkersOnly) +
                      "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" + XyAxes;
        var bubble = "<c:bubbleChart><c:varyColors val=\"0\"/>" + XySeries(0, "Charts!$B$4:$B$9", "Charts!$C$4:$C$9", "", "Charts!$D$4:$D$9") + "<c:bubbleScale val=\"100\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:bubbleChart>" + XyAxes;
        string RadarAxes() => Axes();
        var radar = "<c:radarChart><c:radarStyle val=\"marker\"/><c:varyColors val=\"0\"/>" + ChartSeries(0, true) + ChartSeries(1, true) + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:radarChart>" + RadarAxes();
        var filled = "<c:radarChart><c:radarStyle val=\"filled\"/><c:varyColors val=\"0\"/>" + ChartSeries(0, true) + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:radarChart>" + RadarAxes();
        string[] charts =
        [
            ChartSpace("Stacked areas", area, "b"), ChartSpace("Doughnut shares", doughnut), ChartSpace("Rings", rings, "b"),
            ChartSpace("Scatter", scatter, "b"), ChartSpace("Bubbles", bubble, ""), ChartSpace("Radar", radar, "b"), ChartSpace("Filled radar", filled, ""),
        ];
        (int C0, int R0, int C1, int R1)[] places = [(0, 2, 3, 15), (3, 2, 6, 15), (0, 16, 3, 29), (3, 16, 6, 29), (0, 30, 3, 43), (3, 30, 6, 43), (0, 44, 3, 57)];
        WriteChartParts(zip, sheetNumber, charts, places, "more");
    }

    private static string MoreChartsSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"6\" width=\"16\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "More charts: area, doughnut, scatter, bubble and radar (data on the Charts sheet)", Bold)),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"1\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    private static void WriteChartOptions(ZipArchive zip, int sheetNumber)
    {
        string Series(int i, string tail = "") => ChartSeries(i, true).Replace("</c:ser>", tail + "</c:ser>");
        var secondaryAxes = Axes("", "Units") +
            "<c:catAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\"/><c:axPos val=\"b\"/><c:crossAx val=\"4\"/></c:catAx>" +
            "<c:valAx><c:axId val=\"4\"/><c:scaling><c:orientation val=\"minMax\"/><c:max val=\"40\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/>" + Rich("title", "Rate", 1000) +
            "<c:numFmt formatCode=\"0&quot; %&quot;\" sourceLinked=\"0\"/><c:crossAx val=\"3\"/><c:crosses val=\"max\"/></c:valAx>";
        var combo = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + Series(0) + "<c:gapWidth val=\"150\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                    "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + Series(1) + "<c:marker val=\"1\"/><c:axId val=\"3\"/><c:axId val=\"4\"/></c:lineChart>" + secondaryAxes;
        var smooth = "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + Series(0, "<c:smooth val=\"1\"/>") + Series(1, "<c:smooth val=\"0\"/>") + "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" + Axes();
        var reversed = Bars("bar", "clustered", 1, true, axes: Axes().Replace("<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/>", "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"maxMin\"/>"));
        var log = Bars("col", "clustered", 2, true, axes: Axes().Replace("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/>", "<c:valAx><c:axId val=\"2\"/><c:scaling><c:logBase val=\"10\"/><c:orientation val=\"minMax\"/>"));
        var flipped = Bars("col", "clustered", 1, true, axes: Axes().Replace("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/>", "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"maxMin\"/>"));
        // Two scatter groups: the second on its own pair of value axes (a secondary Y axis on the right, a secondary X axis on top).
        var secondaryXy = "<c:valAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/><c:max val=\"20\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"t\"/>" + Rich("title", "Secondary X", 1000) + "<c:crossAx val=\"4\"/></c:valAx>" +
                          "<c:valAx><c:axId val=\"4\"/><c:scaling><c:orientation val=\"minMax\"/><c:max val=\"5000\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/>" + Rich("title", "Secondary Y", 1000) + "<c:crossAx val=\"3\"/><c:crosses val=\"max\"/></c:valAx>";
        var scatterPair = "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>" + XySeries(0, "Charts!$B$4:$B$9", "Charts!$C$4:$C$9") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" +
                          "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>" + XySeries(1, "Charts!$B$4:$B$9", "Charts!$D$4:$D$9") + "<c:axId val=\"3\"/><c:axId val=\"4\"/></c:scatterChart>" + XyAxes + secondaryXy;
        string[] charts =
        [
            ChartSpace("Columns and a line on a second axis", combo, "b"), ChartSpace("Smoothed and straight lines", smooth, "b"),
            ChartSpace("Categories reversed", reversed, ""), ChartSpace("Logarithmic axis", log, "b"), ChartSpace("Values reversed", flipped, ""),
            ChartSpace("Scatter on secondary axes", scatterPair, "b"),
        ];
        (int C0, int R0, int C1, int R1)[] places = [(0, 2, 3, 15), (3, 2, 6, 15), (0, 16, 3, 29), (3, 16, 6, 29), (0, 30, 3, 43), (3, 30, 6, 43)];
        WriteChartParts(zip, sheetNumber, charts, places, "option");
    }

    private static string ChartOptionsSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"6\" width=\"16\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "Chart options: secondary axis, smoothing, reversed and logarithmic axes", Bold)),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"1\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    private static string Alpha(string hex, int percent) =>
        $"<c:spPr><a:solidFill><a:srgbClr val=\"{hex}\"><a:alpha val=\"{percent * 1000}\"/></a:srgbClr></a:solidFill></c:spPr>";

    private const string FixedBars = "<c:errBars><c:errDir val=\"y\"/><c:errBarType val=\"both\"/><c:errValType val=\"fixedVal\"/><c:noEndCap val=\"0\"/><c:val val=\"4\"/>" +
                                     "<c:spPr><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"404040\"/></a:solidFill></a:ln></c:spPr></c:errBars>";

    private static string Trend(string type, string extra = "", bool labels = false) =>
        $"<c:trendline><c:spPr><a:ln w=\"22225\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></c:spPr><c:trendlineType val=\"{type}\"/>{extra}<c:dispRSqr val=\"{(labels ? 1 : 0)}\"/><c:dispEq val=\"{(labels ? 1 : 0)}\"/></c:trendline>";

    private static void WriteChartExtras(ZipArchive zip, int sheetNumber)
    {
        string Series(int i, string tail = "") => ChartSeries(i, true).Replace("<c:cat>", tail + "<c:cat>");
        string ManualLayout(string target, double x, double y, double w, double h) =>
            $"<c:layout><c:manualLayout><c:layoutTarget val=\"{target}\"/><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"{x}\"/><c:y val=\"{y}\"/><c:w val=\"{w}\"/><c:h val=\"{h}\"/></c:manualLayout></c:layout>";
        var areas = "<c:areaChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + Series(0, Alpha("4472C4", 55)) + Series(1, Alpha("ED7D31", 55)) + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>" +
                    Axes().Replace("<c:crossAx val=\"1\"/></c:valAx>", "<c:crossAx val=\"1\"/><c:crossBetween val=\"midCat\"/></c:valAx>");
        var bubbles = "<c:bubbleChart><c:varyColors val=\"0\"/>" + XySeries(0, "Charts!$B$4:$B$9", "Charts!$C$4:$C$9", Alpha("5B9BD5", 50), "Charts!$D$4:$D$9") + "<c:bubbleScale val=\"120\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:bubbleChart>" + XyAxes;
        var trend = "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + Series(0, Trend("linear", "<c:forward val=\"1\"/>") + FixedBars) + "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" + Axes();
        var scatter = "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>" +
                      XySeries(0, "Charts!$B$4:$B$9", "Charts!$C$4:$C$9", "<c:spPr><a:ln w=\"19050\"><a:noFill/></a:ln></c:spPr>" + Trend("poly", "<c:order val=\"2\"/>")) + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" + XyAxes;
        var columns = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + Series(0, FixedBars.Replace("fixedVal", "percentage").Replace("<c:val val=\"4\"/>", "<c:val val=\"15\"/>")) +
                      "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes();
        var manual = ChartSpace("Manual plot layout", columns, "").Replace("<c:plotArea><c:layout/>", "<c:plotArea>" + ManualLayout("inner", 0.3, 0.25, 0.6, 0.5));
        string[] charts =
        [
            ChartSpace("Translucent areas", areas, "b"), ChartSpace("Translucent bubbles", bubbles, ""), ChartSpace("Trendline and error bars", trend, "b"),
            ChartSpace("Polynomial trendline", scatter, ""), manual,
        ];
        (int C0, int R0, int C1, int R1)[] places = [(0, 2, 3, 15), (3, 2, 6, 15), (0, 16, 3, 29), (3, 16, 6, 29), (0, 30, 3, 43)];
        WriteChartParts(zip, sheetNumber, charts, places, "extra");
    }

    private static void WriteChartText(ZipArchive zip, int sheetNumber)
    {
        string Series(int i, string tail = "") => ChartSeries(i, true).Replace("<c:cat>", tail + "<c:cat>");
        string Fill(string hex) => $"<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>";
        string TextProperties(string attributes, string hex = "", string face = "") =>
            $"<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr {attributes}>{(hex.Length == 0 ? "" : Fill(hex))}{(face.Length == 0 ? "" : $"<a:latin typeface=\"{face}\"/>")}</a:defRPr></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>";
        var columns = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + Series(0, Trend("linear", "", labels: true)) +
                      "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes();
        var bars = "<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + Series(1, Trend("linear")) +
                   "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes();
        var stacked = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"stacked\"/><c:varyColors val=\"0\"/>" + Series(0) + Series(1, Trend("linear", "", labels: true)) +
                      "<c:gapWidth val=\"80\"/><c:overlap val=\"100\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes();
        var scatter = "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>" +
                      XySeries(0, "Charts!$B$4:$B$9", "Charts!$C$4:$C$9", "<c:spPr><a:ln w=\"19050\"><a:noFill/></a:ln></c:spPr>" + Trend("poly", "<c:order val=\"2\"/>", labels: true)) + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" + XyAxes;
        // Text formatting: a colored bold title, axis labels in other sizes, colors and fonts, and a legend placed by hand.
        var styled = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + Series(0, "<c:dLbls>" + TextProperties("sz=\"900\" b=\"1\"", "C00000") + "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbls>") + Series(1) +
                     "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                     Axes("Quarter", "Units").Replace("<c:crossAx val=\"2\"/></c:catAx>", TextProperties("sz=\"1100\" i=\"1\"", "548235", "Georgia") + "<c:crossAx val=\"2\"/></c:catAx>")
                                             .Replace("<c:crossAx val=\"1\"/></c:valAx>", TextProperties("sz=\"800\"", "2F5597") + "<c:crossAx val=\"1\"/></c:valAx>");
        var styledSpace = ChartSpace("Formatted text", styled, "")
            .Replace("<a:rPr lang=\"en-US\" sz=\"1400\"/><a:t>Formatted text", "<a:rPr lang=\"en-US\" sz=\"1800\" b=\"1\">" + Fill("1F3864") + "<a:latin typeface=\"Georgia\"/></a:rPr><a:t>Formatted text")
            .Replace("</c:plotArea>", "</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.62\"/><c:y val=\"0.2\"/><c:w val=\"0.2\"/><c:h val=\"0.2\"/></c:manualLayout></c:layout><c:overlay val=\"1\"/>" +
                     TextProperties("sz=\"1000\" b=\"1\"", "7030A0") + "</c:legend>");
        // Axis labels turned 90° (rot is clockwise in the file), and labels on single points: one deleted, one with its own format and style.
        string PointLabels(string inner) => "<c:dLbls>" + inner + "<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbls>";
        const string Show = "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/>";
        var turned = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" +
                     Series(0, PointLabels("<c:dLbl><c:idx val=\"2\"/><c:numFmt formatCode=\"0.0&quot; units&quot;\" sourceLinked=\"0\"/>" + TextProperties("sz=\"1100\" b=\"1\"", "C00000") + Show + "</c:dLbl>" +
                                                "<c:dLbl><c:idx val=\"4\"/>" + Show + "</c:dLbl>")) +
                     "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                     Axes().Replace("<c:crossAx val=\"2\"/></c:catAx>", "<c:txPr><a:bodyPr rot=\"-5400000\" vert=\"horz\"/><a:p><a:pPr><a:defRPr sz=\"900\"/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr><c:crossAx val=\"2\"/></c:catAx>");
        // A trendline label with its own number format, moved from its place, and a legend given a size (the entries flow into columns).
        var trendLabel = "<c:trendline><c:spPr><a:ln w=\"22225\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></c:spPr><c:trendlineType val=\"linear\"/><c:dispRSqr val=\"1\"/><c:dispEq val=\"1\"/>" +
                         "<c:trendlineLbl><c:layout><c:manualLayout><c:x val=\"-0.25\"/><c:y val=\"0.12\"/></c:manualLayout></c:layout><c:numFmt formatCode=\"0.000\" sourceLinked=\"0\"/></c:trendlineLbl></c:trendline>";
        var formatted = "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + Series(0, trendLabel) + Series(1) + Series(2) + "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" + Axes();
        var formattedSpace = ChartSpace("Trendline label format, legend size", formatted, "")
            .Replace("</c:plotArea>", "</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.55\"/><c:y val=\"0.12\"/><c:w val=\"0.4\"/><c:h val=\"0.2\"/></c:manualLayout></c:layout><c:overlay val=\"1\"/></c:legend>");
        // Labels of single points with their own position (inside the base, centered) and text (a field and a literal), under a series position.
        string Turn(string degrees) => $"<c:txPr><a:bodyPr rot=\"{degrees}\" vert=\"horz\"/><a:p><a:pPr><a:defRPr sz=\"900\"/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>";
        var positioned = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" +
                         Series(0, "<c:dLbls><c:dLbl><c:idx val=\"1\"/><c:dLblPos val=\"inBase\"/>" + Show + "</c:dLbl>" +
                                   "<c:dLbl><c:idx val=\"2\"/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Peak </a:t></a:r><a:fld id=\"{1A2B3C4D-0000-0000-0000-000000000001}\" type=\"VALUE\"><a:t>[VALUE]</a:t></a:fld></a:p></c:rich></c:tx>" + Show + "</c:dLbl>" +
                                   "<c:dLbl><c:idx val=\"3\"/><c:dLblPos val=\"ctr\"/>" + Show + "</c:dLbl>" +
                                   "<c:dLblPos val=\"outEnd\"/>" + Show + "</c:dLbls>") +
                         "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes();
        var turnedBars = Bars("bar", "clustered", 1, true, axes: Axes().Replace("<c:crossAx val=\"2\"/></c:catAx>", Turn("-2700000") + "<c:crossAx val=\"2\"/></c:catAx>")
                                                                   .Replace("<c:crossAx val=\"1\"/></c:valAx>", Turn("2700000") + "<c:crossAx val=\"1\"/></c:valAx>"));
        var turnedValues = Bars("col", "clustered", 1, true, axes: Axes().Replace("<c:crossAx val=\"1\"/></c:valAx>", Turn("-5400000") + "<c:crossAx val=\"1\"/></c:valAx>")
                                                                   .Replace("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/>", "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"maxMin\"/>")
                                                                   .Replace("<c:crossAx val=\"2\"/></c:catAx>", Turn("-2700000") + "<c:crossAx val=\"2\"/></c:catAx>"));
        var squeezed = ChartSpace("Legend that shrinks the plot", columns.Replace(Trend("linear", "", labels: true), ""), "")
            .Replace("</c:plotArea>", "</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.7\"/><c:y val=\"0.3\"/><c:w val=\"0.25\"/><c:h val=\"0.3\"/></c:manualLayout></c:layout><c:overlay val=\"0\"/></c:legend>");
        // Pie labels at the rim (outEnd), at best fit and centered, and a label of two paragraphs.
        string PieLabels(string position) => "<c:dLbls><c:dLblPos val=\"" + position + "\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"1\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/><c:separator> </c:separator></c:dLbls>";
        string PieOf(string labels) => "<c:pieChart><c:varyColors val=\"1\"/>" + ChartSeries(0, true).Replace("<c:cat>", labels + "<c:cat>") + "</c:pieChart>";
        var twoLines = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" +
                       Series(0, "<c:dLbls><c:dLbl><c:idx val=\"2\"/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Best quarter</a:t></a:r></a:p><a:p><a:fld id=\"{1A2B3C4D-0000-0000-0000-000000000002}\" type=\"VALUE\"><a:t>[VALUE]</a:t></a:fld></a:p></c:rich></c:tx>" + Show + "</c:dLbl>" +
                                 "<c:dLblPos val=\"outEnd\"/>" + Show + "</c:dLbls>") +
                       "<c:gapWidth val=\"80\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes();
        string[] charts =
        [
            ChartSpace("Columns: equation and R²", columns, ""), ChartSpace("Bars with a trendline", bars, ""), ChartSpace("Stacked columns: trend of the totals", stacked, "b"),
            ChartSpace("Polynomial: equation and R²", scatter, ""), styledSpace,
            ChartSpace("Turned labels and point labels", turned, ""), formattedSpace,
            ChartSpace("Label positions and own text", positioned, ""), ChartSpace("Turned bar labels", turnedBars, ""),
            ChartSpace("Turned labels, reversed values", turnedValues, ""), squeezed,
            ChartSpace("Pie labels at the rim", PieOf(PieLabels("outEnd")), "r"), ChartSpace("Pie labels at best fit", PieOf(PieLabels("bestFit")), "r"), ChartSpace("Two-line label", twoLines, ""),
        ];
        (int C0, int R0, int C1, int R1)[] places = [(0, 2, 3, 15), (3, 2, 6, 15), (0, 16, 3, 29), (3, 16, 6, 29), (0, 30, 3, 43), (3, 30, 6, 43), (0, 44, 3, 57),
            (3, 44, 6, 57), (0, 58, 3, 71), (3, 58, 6, 71), (0, 72, 3, 85), (3, 72, 6, 85), (0, 86, 3, 99), (3, 86, 6, 99)];
        WriteChartParts(zip, sheetNumber, charts, places, "text");
    }

    private static string ChartTextSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"6\" width=\"16\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "Chart text: trendlines on bars and stacks, equations, text styles and a manual legend (data on the Charts sheet)", Bold)),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"1\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    private static string ChartExtrasSheet() => Sheet("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"6\" width=\"16\" customWidth=\"1\"/></cols>",
        Row(1, Str("A1", "Chart extras: transparency, trendlines, error bars and manual layout (data on the Charts sheet)", Bold)),
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"1\"/>" + HeaderFooter + "<drawing r:id=\"rId1\"/>");

    // ------------------------------------------------------------------ comments

    private static void WriteComments(ZipArchive zip, int sheetNumber, StringBuilder content)
    {
        var longText = string.Join(" ", Enumerable.Repeat("This note is deliberately long so that it wraps over several lines of the list.", 4));
        Entry(zip, $"xl/comments{sheetNumber}.xml",
            "<comments xmlns=\"" + Main + "\"><authors><author>Ada</author><author>Grace</author></authors><commentList>" +
            "<comment ref=\"B4\" authorId=\"1\"><text><r><rPr><b/><sz val=\"9\"/></rPr><t>Grace:</t></r><r><t xml:space=\"preserve\">\nCheck this figure against the ledger.</t></r></text></comment>" +
            "<comment ref=\"A3\" authorId=\"0\"><text><t>Starting balance, carried over.</t></text></comment>" +
            $"<comment ref=\"C6\" authorId=\"0\"><text><t>{longText}</t></text></comment></commentList></comments>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}comments\" Target=\"../comments{sheetNumber}.xml\"/></Relationships>");
        content.Append($"<Override PartName=\"/xl/comments{sheetNumber}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>");
    }

    private static string CommentsSheet()
    {
        var rows = Row(1, Str("A1", "Cell comments", Bold)) + Row(3, Str("A3", "Opening") + Num("B3", 1200)) + Row(4, Str("A4", "Adjustment") + Num("B4", -75.5)) +
                   Row(6, Str("A6", "Total") + Num("B6", 1124.5) + Str("C6", "see note"));
        return Sheet("", rows, "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" cellComments=\"atEnd\"/>" + HeaderFooter);
    }

    private static string AreasSheet()
    {
        var rows = Row(1, Str("A1", "Print area one", Bold) + Str("E1", "never printed")) +
                   Row(3, Str("C3", "End one")) +
                   Row(5, Str("E5", "Print area two", Bold)) + Row(7, Str("F7", "End two")) +
                   Row(10, Str("A10", "Print area three", Bold)) + Row(11, Str("B11", "End three")) +
                   Row(14, Str("A14", "never printed"));
        return Sheet("", rows, "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    private static string PaperSheet()
    {
        var rows = Row(1, Str("A1", "Paper sizes", Bold)) + Row(3, Str("A3", "This sheet is printed on A3 paper (paperSize 8)."));
        return Sheet("", rows, "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"8\"/>" + HeaderFooter);
    }

    /// <summary>A6 paper (paperSize 70), page numbering restarted at 20, and fill alignment repeating a text across its cell.</summary>
    /// <summary>The print area starts at row 5; row 1 lies outside it and is printed on every page anyway. The page is scaled to 60 % and so is its footer.</summary>
    private static string TitlesSheet()
    {
        var rows = new StringBuilder(Row(1, Str("A1", "Print title outside the area", Bold) + Str("B1", "Repeats on every page", Bold)));
        rows.Append(Row(2, Str("A2", "Rows 2 to 4 are not printed")));
        for (int r = 5; r <= 124; r++) rows.Append(Row(r, Str($"A{r}", $"Titled row {r}") + Str($"B{r}", $"value {r * 3}")));
        return Sheet("<cols><col min=\"1\" max=\"2\" width=\"28\" customWidth=\"1\"/></cols>", rows.ToString(),
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" scale=\"60\"/>" + HeaderFooter);
    }

    private static string NumberingSheet()
    {
        var rows = Row(1, Str("A1", "Numbering, A6 paper and fill alignment", Bold)) + Row(3, Str("A3", "-=", FillAlign)) + Row(4, Str("A4", "ab", FillAlign));
        return Sheet("<cols><col min=\"1\" max=\"1\" width=\"30\" customWidth=\"1\"/></cols>", rows,
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"70\" firstPageNumber=\"20\" useFirstPageNumber=\"1\"/>" +
            "<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddFooter>&amp;CCounted &amp;P of &amp;N</oddFooter></headerFooter>");
    }

    private static string GradientSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Gradient fills", Bold)));
        rows.Append(Row(3, Str("A3", "Left to right", GradLinear) + Str("B3", "Diagonal, three stops (theme, rgb, indexed)", GradDiagonal), 54));
        rows.Append(Row(5, Str("A5", "Top to bottom", GradVertical) + Str("B5", "Path from the center", GradCenter), 54));
        rows.Append(Row(7, Str("A7", "Path from a corner", GradCorner), 54));
        var cols = "<cols><col min=\"1\" max=\"1\" width=\"26\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"34\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    // ------------------------------------------------------------------ pictures

    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string DmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const long Emu = 12700;

    private static string Pic(string name, string extra = "", int rotation = 0, string flip = "", string embed = "rId1") =>
        $"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"{name}\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"{embed}\"/>{extra}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>" +
        $"<xdr:spPr><a:xfrm rot=\"{rotation}\"{flip}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>";

    private static string Marker(string tag, int col, int row, double colOffset = 0, double rowOffset = 0) =>
        $"<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>{(long)(colOffset * Emu)}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{(long)(rowOffset * Emu)}</xdr:rowOff></xdr:{tag}>";

    private static string TwoCell(int c0, int r0, int c1, int r1, string content) =>
        $"<xdr:twoCellAnchor>{Marker("from", c0, r0)}{Marker("to", c1, r1)}{content}<xdr:clientData/></xdr:twoCellAnchor>";

    private static string OneCell(int col, int row, double colOffset, double rowOffset, double width, double height, string content) =>
        $"<xdr:oneCellAnchor>{Marker("from", col, row, colOffset, rowOffset)}<xdr:ext cx=\"{(long)(width * Emu)}\" cy=\"{(long)(height * Emu)}\"/>{content}<xdr:clientData/></xdr:oneCellAnchor>";

    private static string Absolute(double x, double y, double width, double height, string content) =>
        $"<xdr:absoluteAnchor><xdr:pos x=\"{(long)(x * Emu)}\" y=\"{(long)(y * Emu)}\"/><xdr:ext cx=\"{(long)(width * Emu)}\" cy=\"{(long)(height * Emu)}\"/>{content}<xdr:clientData/></xdr:absoluteAnchor>";

    private static void WritePictures(ZipArchive zip, int sheetNumber, StringBuilder content)
    {
        var anchors =
            TwoCell(1, 3, 4, 8, Pic("Stretched to B4:D8")) +                                                                       // stretches with its cells
            OneCell(5, 3, 8, 6, 120, 75, Pic("One-cell anchor")) +                                                                // fixed size
            OneCell(5, 9, 4, 4, 120, 75, Pic("Cropped", "<a:srcRect l=\"25000\" r=\"25000\" t=\"10000\"/>")) +                          // crop
            OneCell(1, 10, 10, 10, 120, 75, Pic("Rotated", rotation: 20 * 60000)) +                                               // rotation
            OneCell(5, 21, 4, 2, 56, 35, Pic("Flipped horizontally", flip: " flipH=\"1\"")) +                                    // mirrored left to right
            OneCell(5, 24, 4, 2, 56, 35, Pic("Flipped vertically", flip: " flipV=\"1\"")) +                                      // mirrored top to bottom
            Absolute(300, 260, 90, 56, Pic("Absolute")) +                                                                         // measured from A1
            OneCell(1, 28, 4, 2, 90, 56, Pic("Bitmap", embed: "rId2")) +                                                          // BMP
            OneCell(4, 28, 4, 2, 90, 56, Pic("Animated-format GIF", embed: "rId3")) +                                            // GIF with a transparent corner
            OneCell(1, 33, 4, 2, 90, 56, Pic("RLE-compressed bitmap", embed: "rId4")) +                                           // BMP, RLE8
            OneCell(4, 33, 4, 2, 90, 56, Pic("LZW TIFF", embed: "rId5")) +                                                      // TIFF, LZW with the horizontal predictor
            OneCell(4, 38, 4, 2, 90, 56, Pic("Three-strip JPEG TIFF", embed: "rId7")) +                                           // TIFF with three JPEG strips (written by libtiff), joined with restart markers
            OneCell(10, 38, 4, 2, 90, 56, Pic("Fax TIFF", embed: "rId9")) +                                                      // CCITT Group 4, passed to the PDF as it is (CCITTFaxDecode)
            OneCell(7, 38, 4, 2, 90, 56, Pic("Tiled JPEG TIFF", embed: "rId8")) +                                                // TIFF of twelve 16x16 JPEG tiles (written by libtiff), re-coded into one JPEG
            OneCell(1, 38, 4, 2, 90, 56, Pic("JPEG TIFF", embed: "rId6")) +                                                       // TIFF holding one JPEG strip (compression 7) and JPEGTables
            TwoCell(1, 17, 4, 24, Pic("Across the page break")) +                                                                  // continues on the next page
            "<xdr:oneCellAnchor>" + Marker("from", 6, 17) + "<xdr:ext cx=\"1270000\" cy=\"635000\"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"9\" name=\"Freeform\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:custGeom/></xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>";
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/drawings/_rels/drawing{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}image\" Target=\"../media/image1.png\"/><Relationship Id=\"rId2\" Type=\"{RelType}image\" Target=\"../media/image2.bmp\"/><Relationship Id=\"rId3\" Type=\"{RelType}image\" Target=\"../media/image3.gif\"/><Relationship Id=\"rId4\" Type=\"{RelType}image\" Target=\"../media/image4.bmp\"/><Relationship Id=\"rId5\" Type=\"{RelType}image\" Target=\"../media/image5.tif\"/><Relationship Id=\"rId6\" Type=\"{RelType}image\" Target=\"../media/image6.tif\"/><Relationship Id=\"rId7\" Type=\"{RelType}image\" Target=\"../media/image7.tif\"/><Relationship Id=\"rId8\" Type=\"{RelType}image\" Target=\"../media/image8.tif\"/><Relationship Id=\"rId9\" Type=\"{RelType}image\" Target=\"../media/image9.tif\"/></Relationships>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/><Relationship Id=\"rId2\" Type=\"{RelType}vmlDrawing\" Target=\"../drawings/vmlDrawing1.vml\"/></Relationships>");
        // The pictures of the header and footer (&G) come from a legacy VML part.
        Entry(zip, "xl/drawings/vmlDrawing1.vml",
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">" +
            "<v:shape id=\"LH\" o:spid=\"_x0000_s1025\" type=\"#_x0000_t75\" style=\"position:absolute;margin-left:0;margin-top:0;width:54pt;height:30pt;z-index:1\"><v:imagedata o:relid=\"rId1\" o:title=\"logo\"/><o:lock v:ext=\"edit\" rotation=\"t\"/></v:shape>" +
            "<v:shape id=\"LHFIRST\" o:spid=\"_x0000_s1027\" type=\"#_x0000_t75\" style=\"position:absolute;margin-left:0;margin-top:0;width:80pt;height:24pt;z-index:3\"><v:imagedata o:relid=\"rId1\" o:title=\"logo\"/></v:shape>" +
            "<v:shape id=\"RF\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t75\" style=\"position:absolute;margin-left:0;margin-top:0;width:36pt;height:36pt;z-index:2\"><v:imagedata o:relid=\"rId1\" o:title=\"logo\"/></v:shape></xml>");
        Entry(zip, "xl/drawings/_rels/vmlDrawing1.vml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}image\" Target=\"../media/image1.png\"/></Relationships>");
        content.Append("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");
        using (var media = zip.CreateEntry("xl/media/image1.png").Open()) media.Write(ShowcaseDocx.GradientPng(160, 100));
        using (var media = zip.CreateEntry("xl/media/image2.bmp").Open()) media.Write(GradientBmp(90, 56));
        using (var media = zip.CreateEntry("xl/media/image3.gif").Open()) media.Write(CheckerGif(48, 32));
        using (var media = zip.CreateEntry("xl/media/image4.bmp").Open()) media.Write(RleBmp(72, 40));
        using (var media = zip.CreateEntry("xl/media/image5.tif").Open()) media.Write(LzwTiff(64, 40));
        using (var media = zip.CreateEntry("xl/media/image6.tif").Open()) media.Write(JpegTiff());
        using (var media = zip.CreateEntry("xl/media/image7.tif").Open()) media.Write(StripedJpegTiff());
        using (var media = zip.CreateEntry("xl/media/image8.tif").Open()) media.Write(TiledJpegTiff());
        using (var media = zip.CreateEntry("xl/media/image9.tif").Open()) media.Write(FaxTiff());
        content.Append("<Default Extension=\"tif\" ContentType=\"image/tiff\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"bmp\" ContentType=\"image/bmp\"/><Default Extension=\"gif\" ContentType=\"image/gif\"/>");
    }

    /// <summary>A 24-bit bottom-up BMP: red grows to the right, blue towards the top.</summary>
    private static byte[] GradientBmp(int width, int height)
    {
        int stride = (width * 3 + 3) / 4 * 4;
        var ms = new MemoryStream();
        var w = new BinaryWriter(ms);
        w.Write("BM"u8.ToArray()); w.Write(54 + stride * height); w.Write(0); w.Write(54);
        w.Write(40); w.Write(width); w.Write(height); w.Write((short)1); w.Write((short)24); w.Write(0); w.Write(stride * height); w.Write(0); w.Write(0); w.Write(0); w.Write(0);
        for (int y = height - 1; y >= 0; y--)
        {
            var row = new byte[stride];
            for (int x = 0; x < width; x++)
            {
                row[x * 3] = (byte)(255 * y / height);        // blue
                row[x * 3 + 1] = 90;                           // green
                row[x * 3 + 2] = (byte)(255 * x / width);     // red
            }
            // BMP rows run bottom-up: the first row written is the bottom of the picture.
            w.Write(row);
        }
        return ms.ToArray();
    }

    /// <summary>An 8-bit RLE-compressed BMP: four diagonal color bands, one run per band and row (rows run bottom-up).</summary>
    private static byte[] RleBmp(int width, int height)
    {
        var data = new List<byte>();
        for (int y = height - 1; y >= 0; y--)
        {
            int x = 0;
            while (x < width)
            {
                int band = (x + y) / 18 % 4, run = 1;
                while (x + run < width && (x + run + y) / 18 % 4 == band && run < 255) run++;
                data.Add((byte)run); data.Add((byte)band);
                x += run;
            }
            data.Add(0); data.Add(y == 0 ? (byte)1 : (byte)0);   // end of line; end of bitmap after the last row
        }
        var ms = new MemoryStream();
        var w = new BinaryWriter(ms);
        int offset = 54 + 16;
        w.Write("BM"u8.ToArray()); w.Write(offset + data.Count); w.Write(0); w.Write(offset);
        w.Write(40); w.Write(width); w.Write(height); w.Write((short)1); w.Write((short)8); w.Write(1); w.Write(data.Count); w.Write(0); w.Write(0); w.Write(4); w.Write(0);
        w.Write(new byte[] { 40, 40, 200, 0, 40, 160, 40, 0, 200, 80, 40, 0, 40, 200, 240, 0 }); // BGRA
        w.Write(data.ToArray());
        return ms.ToArray();
    }

    /// <summary>
    /// A TIFF whose one strip is a JPEG (compression 7, YCbCr), the way libtiff writes it: the tables in JPEGTables, the frame and scan in the strip.
    /// The JPEG is a 32x20 gradient.
    /// </summary>
    private static byte[] JpegTiff()
    {
        var jpeg = System.Convert.FromBase64String("/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAoHBwgHBgoICAgLCgoLDhgQDg0NDh0VFhEYIx8lJCIfIiEmKzcvJik0KSEiMEExNDk7Pj4+JS5ESUM8SDc9Pjv/2wBDAQoLCw4NDhwQEBw7KCIoOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozv/wAARCAAUACADAREAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDhItN/2a+6rVzgoYgtxab/ALNeNWrnuUMQXItN/wBmvHrVz3aGILkWm/7NeNWrnuUMQWotN/2a+xrVz8KoYgtxab/s141aue5QxBci03/Zrxq1c92hiC5Fpv8As149aue5QxA+KCP0r7StNn4fQnItxQR+lePWmz3KE5FyKCP0rxq02e5QnIuRQR+leNWmz3aE5H//2Q==");
        int frame = Enumerable.Range(2, jpeg.Length - 3).First(i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xC0);
        byte[] tables = [.. jpeg[..frame], 0xFF, 0xD9];
        byte[] strip = [0xFF, 0xD8, .. jpeg[frame..]];
        (int Tag, int Type, int Count, int Value)[] entries =
            [(256, 3, 1, 32), (257, 3, 1, 20), (258, 3, 1, 8), (259, 3, 1, 7), (262, 3, 1, 6), (273, 4, 1, 8), (277, 3, 1, 3), (278, 3, 1, 20), (279, 4, 1, strip.Length), (347, 7, tables.Length, 8 + strip.Length + 2 + 10 * 12 + 4)];
        var ms = new MemoryStream();
        void W16(int v) { ms.WriteByte((byte)v); ms.WriteByte((byte)(v >> 8)); }
        void W32(int v) { W16(v & 0xFFFF); W16(v >> 16); }
        ms.Write("II"u8); W16(42); W32(8 + strip.Length);
        ms.Write(strip);
        W16(entries.Length);
        foreach (var (tag, type, count, value) in entries)
        {
            W16(tag); W16(type); W32(count);
            if (type == 3) { W16(value); W16(0); } else W32(value);
        }
        W32(0);
        ms.Write(tables);
        return ms.ToArray();
    }

    /// <summary>A 40x48 TIFF of three JPEG strips of 16 rows (YCbCr 2x2, shared JPEGTables), as libtiff writes it (tiffcp -c jpeg -r 16).</summary>
    private static byte[] StripedJpegTiff() => System.Convert.FromBase64String("SUkqAE4BAAD/2P/AABEIABAAKAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APEoLXpxWjBa9OKswWvTitGC16cVUWXhMUVoLXpxWjBa9OKswWvTitGC16cVvFn0+ExRWgtenFFbcFr04orZM+ipYr3T/9n/2P/AABEIABAAKAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/AOGgtenFaUFr04qzBa9OK0YLXpxRFng4TElaC16cVowWvTirMFr7VowWvTit4s+nwmJK0Fr04oragtenFFbJn0NLE+6f/9n/2P/AABEIABAAKAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/AI4LXpxWjBa9OKswWvTitGC16cVMWfn+ExRWgtenFaMFr04qzBa9OK0YLXpxW8WfUYTElaC16cUVtQ2wAyeAO5oqpVqdPSckvVn0VHEtxP/ZAA4AAAEDAAEAAAAoAAAAAQEDAAEAAAAwAAAAAgEDAAMAAAD8AQAAAwEDAAEAAAAHAAAABgEDAAEAAAAGAAAAEQEEAAMAAAAIAgAAEgEDAAEAAAABAAAAFQEDAAEAAAADAAAAFgEDAAEAAAAQAAAAFwEDAAMAAAACAgAAHAEDAAEAAAABAAAAWwEHAD4CAABEAgAAEgIDAAIAAAACAAIAFAIFAAYAAAAUAgAAAAAAAAgACAAIAGkAaQBzAAgAAABxAAAA2gAAAAAAAAABAAAA/wAAAAEAAACAAAAAAQAAAP8AAAABAAAAgAAAAAEAAAD/AAAAAQAAAP/Y/9sAQwAIBgYHBgUIBwcHCQkICgwUDQwLCwwZEhMPFB0aHx4dGhwcICQuJyAiLCMcHCg3KSwwMTQ0NB8nOT04MjwuMzQy/9sAQwEJCQkMCwwYDQ0YMiEcITIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIy/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9k=");

    /// <summary>A 64x48 TIFF of twelve 16x16 JPEG tiles (YCbCr 2x2, shared JPEGTables), as libtiff writes it (tiffcp -c jpeg:90 -t -w 16 -l 16).</summary>
    private static byte[] TiledJpegTiff() => System.Convert.FromBase64String("SUkqAJAJAAD/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APhSz0aKdQsEu9HjbhgOuPvHn5cgk4AHDE4443YfDRzb/ardfLbDMq8KgYdcjpg5AIPY+ox0ug6MJrmCWV0LoAXLEZJ7kg4HTHQ9uORkdjZ+GFmSODIiTKo6ISeMkEqoGVOCDg4zk+hzlkWV/wBqOKte58/QziUGlzaH/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APknS/CK7S8YaHOWVSAwySCxJGBxnbz2xzgk1rWPgSZrj7QY93llyXjQ/Ki9MdcDv0J575rvm0620p1hiiaWfbwSFieMD/aPfGTjGSO/etCKx1JD5cs80IEiskUOeDgZBY/MRjo2eN3oBX19fLsjwFZ4Vwdats4wSfL5Sk7JeaV2nukexgc2rTs4ytFn/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APnu28HpIIY5SFhHLTFyGA+8WJYdyM9uM8d6t2nhixVWSa7CTMATHGHkZGHGDwRzknB5/EiutsPBdjpe2S4uY4IrgGVFCjJxznABPAAOR02sB1rc07RWikEdnZSSzoFaM3Shfn285QcDhlOTzj07/Q16WAoP2csLSpy/lk5VZed4wUWr+aXz1P1bA5rNLSd7fLbzfnvY/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APGLXw7oxNuFWaXaS0f2ZMkEDqSwVsZzxgdG56CtuLwnYeW6/Y7gpkOIHEfKgZDY6DqowcgDt6dfaabLHOm+a2skVnf5YyjSKc8gMTnHXHQg46Vq6ZYsLsJbz3bs+5mKFgvGTluF+XrgjH3hzlcD06ywdKvyzoQi9HyuEY6O2qUpyqPzuvPY/pbA5tUSXLLbzf8AV/mf/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APJND0QQWzvEPMPlgvOse7ZgAHdz1ABGMY5PB5x0Mdi2n2UMg3iWZv3Z2ByhHXqOo4xgjqDyM46XQNAtI4bUnIZcyLhsPj5cMM5zkjOQMHrjOK2odFa88TXEUCYiiYRp5aMhAGQ3Gemcj0+bPHQRkWZVcBkMZYW6rVmoRfWP80vklZW2bTWx/PeEzPnqOU9Ulf8AT+vmf//Z/9j/wAARCAAQABADASIAAhEBAxEB/9oADAMBAAIRAxEAPwDktD8KL5EisqPIgBEgwVK8ElcnqMMenP0yDs6faS3c0dpbRLb23K+d5ZVnAJXGM5Byf97PcbiB2dxphuJ4LeLesSMC7IoJmcMuFOO2SB35yOoq0dFjV5beyj8uaIMyyHIaJeRwWPfJOc8HA64qXjaWDovCYVunTpvllKPx1J7ckOyT3fq21Fa/nOBzRzfNJ3v+H9dD/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/AIrXw5ZaNAoabzbl1UFCeAGwV3An6dc/KT+O6NJadEuDdLZ2zFpfvESc9VVuGbIXH4fgOli8MjSXt44LeSW82hmMa7vJXAxnkhshTx9OuAK0bbQo7e+WS9aS9vEUSLGH5GSDyfu8AEYGAMgDrx5EsbLDxdKKVOMdHGLajG6+GUkuepK28IWvquVPQ58JmkZtPmu+/wCvZLzey79f/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/AOm0/RUihlktbX/SBNy7j5dxH8ODnjB6kdMdcKNJtKtZLKF0YeW0m5YYpBgAgnaCDwASw6jsc4Oa6uxsFkfyZ41dzGFa3SMKpG0sQwB6fKpyxJ5yB1FattossF0kbI0PkpnZFnzCSQAuScfKGGOmemDwK+QqY6OGcoRbgo2Vl7i8lyxaWvTnnOevwaJH3OCzTl0bv8/Lz9O33H//2f/Y/8AAEQgAEAAQAwEiAAIRAQMRAf/aAAwDAQACEQMRAD8A6fw9pqwY+1W8qxxx+Zvj2KwUABQccA8e2O+ccbOmaCbC0ebepSIiRWaNTGCR8vQqSDn8vwz8/fD/AOJcmk+HAug37axpL+Xe6rfXUsUt55kn2OVoQHYr5YgMse6ed3Rbe4PmFbTD7nhD45/bvFOlbb+xt/ClysF9faWIYZHsJrplViZkcAqs5WVpGULi6kWQrKI4z5WDyrHUMDh6tHWVGi5R015qslCDtZu94pqL0stWrq/8t4fh/MLznG1l3ve3pq4yleyvp0bV1f8A/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APU9P0L+x9PMdso8wt5cTgDcCW7jgDBIHGcEdByBqWGkzaZbRpgyXEiZZ5WcpFwOvJYNkdDjvx0rxe7+JMOg3/haXw1eanrukacHtEtbG2lttSvQvkSvEloLMW+xLeSIhh5bMhURyL83ma+hfFPU9a+IMVzZi+h0y9kgitZ7vEZuI5LtP9JA8t1jVPNtoEiMazPFeQSyOqqskfgzynF4a0qXuwpxkovZxjGXLKUU/t1Z+7Ta1dlazjdfiGAybMHB1ZLRpvXRuztaz187Wve+mkmv/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APeLLS0ggNnaIJJ1/dySo2XPz/dR/Q5Oevf0Jrc03w/Jas0FtOihE8s3PAZTkDK+gGT19QRgncfAvC37Q934mt9UsNagt/Cmly2dxdtfapFc2yhRHasEhdJt80iRTzPiIo0oEBXyt522NQ+LmsWvhXWtU0lE12DV9auLODS4AoM4kjMVt5VzaKhNyh8qJ4g80jJCskTrGRI3xU8kx9OSpyioSWmj92K0TjGVtEnpOWk5ztCNktPm8v4ezO6hK0W316321+G++l0l1eun/9n/2P/AABEIABAAEAMBIgACEQEDEQH/2gAMAwEAAhEDEQA/APp/TNDbTYIUVDHGc25lZd0sgOPl29jkZ4OeR06HTg8PrahZGtY7SOVQzLuLSyEE9ABwPugHjB78jPknijV9Wg8ZaVol5qFtrFrcG8tb/UNbvEsra0WOPzZBPZJsd9kVrBKzMRFKl1IvyRTx7M23+PmuXvi3wQuqyWHh1Ygtxq0usGSCG/SaQxmazCSP50KKdqyGRlczCTYiRF4/zlZVjVyzhZNRbsrKyu9Ek042cUnFcrTadSc5e4erlmTY7E0Y1aUk7pydm9rStZfE7uLVnG6bj7tnc//ZDwAAAQMAAQAAAEAAAAABAQMAAQAAADAAAAACAQMAAwAAAEoKAAADAQMAAQAAAAcAAAAGAQMAAQAAAAYAAAASAQMAAQAAAAEAAAAVAQMAAQAAAAMAAAAcAQMAAQAAAAEAAABCAQMAAQAAABAAAABDAQMAAQAAABAAAABEAQQADAAAAGgKAABFAQMADAAAAFAKAABbAQcAPgIAAMgKAAASAgMAAgAAAAIAAgAUAgUABgAAAJgKAAAAAAAACAAIAAgAnwCxALoAugC7ALkAwAC/APEA9gDwAPoACAAAAKcAAABYAQAAEgIAAMwCAACHAwAAQAQAAAAFAAC/BQAAsAYAAKYHAACWCAAAAAAAAAEAAAD/AAAAAQAAAIAAAAABAAAA/wAAAAEAAACAAAAAAQAAAP8AAAABAAAA/9j/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/xAAfAAABBQEBAQEBAQAAAAAAAAAAAQIDBAUGBwgJCgv/xAC1EAACAQMDAgQDBQUEBAAAAX0BAgMABBEFEiExQQYTUWEHInEUMoGRoQgjQrHBFVLR8CQzYnKCCQoWFxgZGiUmJygpKjQ1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4eLj5OXm5+jp6vHy8/T19vf4+fr/xAAfAQADAQEBAQEBAQEBAAAAAAAAAQIDBAUGBwgJCgv/xAC1EQACAQIEBAMEBwUEBAABAncAAQIDEQQFITEGEkFRB2FxEyIygQgUQpGhscEJIzNS8BVictEKFiQ04SXxFxgZGiYnKCkqNTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqCg4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2dri4+Tl5ufo6ery8/T19vf4+fr/2Q==");

    /// <summary>A 64x40 bilevel TIFF in CCITT Group 4 (one strip, min-is-black), as libtiff writes it (tiffcp -c g4): a box, a ring, a slash and the word "Fax".</summary>
    private static byte[] FaxTiff() => System.Convert.FromBase64String("SUkqAIAAAAAmoHhv1hZaRBeEF4S8ILyCegvCD8IP0cjCfoJBBv0ggm/SCTf0EE/pIJv7/r8JN+vkGhzx4S/CC/QX4QX5HgSbS5HChf5KEEUP/oEEwu2lwbTVrwQQRQ6tpel0UO2l4QIENLYYS4oRFisIIMLQQYWEEMVCgAgAgAAJAAABAwABAAAAQAAAAAEBAwABAAAAKAAAAAMBAwABAAAABAAAAAYBAwABAAAAAQAAABEBBAABAAAACAAAABIBAwABAAAAAQAAABYBAwABAAAA6AMAABcBBAABAAAAdwAAABwBAwABAAAAAQAAAAAAAAA=");

    /// <summary>An RGB TIFF compressed with LZW (MSB-first codes, early change) and the horizontal predictor.</summary>
    private static byte[] LzwTiff(int width, int height)
    {
        var pixels = new byte[width * height * 3];
        for (int y = 0; y < height; y++)
            for (int x = 0; x < width; x++)
            {
                int o = (y * width + x) * 3;
                pixels[o] = (byte)(255 * x / width); pixels[o + 1] = (byte)(255 * y / height); pixels[o + 2] = (byte)(x / 8 % 2 == y / 8 % 2 ? 220 : 60);
            }
        var diff = (byte[])pixels.Clone();
        for (int y = 0; y < height; y++)
            for (int k = width * 3 - 1; k >= 3; k--) diff[y * width * 3 + k] -= diff[y * width * 3 + k - 3];

        var bits = new List<bool>();
        int size = 9, next = 258;
        void Put(int code) { for (int i = size - 1; i >= 0; i--) bits.Add((code >> i & 1) != 0); }
        var table = new Dictionary<(int, byte), int>();
        Put(256);
        int cur = diff[0];
        for (int i = 1; i < diff.Length; i++)
        {
            if (table.TryGetValue((cur, diff[i]), out var code)) { cur = code; continue; }
            Put(cur);
            table[(cur, diff[i])] = next++;
            if (next > (1 << size) - 1 && size < 12) size++;
            if (next == 4094) { Put(256); size = 9; next = 258; table.Clear(); }
            cur = diff[i];
        }
        Put(cur);
        next++;
        if (next > (1 << size) - 1 && size < 12) size++;
        Put(257);
        var strip = new byte[(bits.Count + 7) / 8];
        for (int i = 0; i < bits.Count; i++) if (bits[i]) strip[i / 8] |= (byte)(0x80 >> i % 8);

        var ms = new MemoryStream();
        var w = new BinaryWriter(ms);
        int ifd = 8 + strip.Length + (strip.Length & 1);
        w.Write("II"u8.ToArray()); w.Write((short)42); w.Write(ifd);
        w.Write(strip); if ((strip.Length & 1) != 0) w.Write((byte)0);
        var entries = new (short Tag, short Type, int Count, int Value)[]
        {
            (256, 3, 1, width), (257, 3, 1, height), (258, 3, 3, ifd + 2 + 11 * 12 + 4), (259, 3, 1, 5), (262, 3, 1, 2),
            (273, 4, 1, 8), (277, 3, 1, 3), (278, 3, 1, height), (279, 4, 1, strip.Length), (284, 3, 1, 1), (317, 3, 1, 2),
        };
        w.Write((short)entries.Length);
        foreach (var (tag, type, count, value) in entries) { w.Write(tag); w.Write(type); w.Write(count); w.Write(value); }
        w.Write(0);
        w.Write(new short[] { 8, 8, 8 }.SelectMany(BitConverter.GetBytes).ToArray());
        return ms.ToArray();
    }

    /// <summary>A GIF with a 2-color checkerboard of 8-pixel squares and a transparent top-left square (index 2), LZW with a clear code every 64 pixels.</summary>
    private static byte[] CheckerGif(int width, int height)
    {
        var ms = new MemoryStream();
        ms.Write("GIF89a"u8);
        ms.Write([(byte)width, 0, (byte)height, 0, 0x80 | 1, 0, 0]);
        ms.Write([200, 40, 40, 40, 80, 200, 255, 255, 255, 0, 0, 0]);
        ms.Write([0x21, 0xF9, 4, 1, 0, 0, 2, 0]);
        ms.Write([0x2C, 0, 0, 0, 0, (byte)width, 0, (byte)height, 0, 0, 2]);
        var bits = new List<bool>();
        void Put(int code) { for (int i = 0; i < 3; i++) bits.Add((code >> i & 1) != 0); }
        for (int i = 0; i < width * height; i++)
        {
            int x = i % width, y = i / width;
            Put(4);                                              // clear before each pixel keeps 3-bit codes
            Put(x < 8 && y < 8 ? 2 : ((x / 8 + y / 8) % 2));
        }
        Put(5);
        var data = new byte[(bits.Count + 7) / 8];
        for (int i = 0; i < bits.Count; i++) if (bits[i]) data[i / 8] |= (byte)(1 << i % 8);
        for (int i = 0; i < data.Length; i += 255)
        {
            int n = Math.Min(255, data.Length - i);
            ms.WriteByte((byte)n); ms.Write(data, i, n);
        }
        ms.Write([0, 0x3B]);
        return ms.ToArray();
    }

    private static string PicturesSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Pictures placed by drawing anchors", Bold)));
        rows.Append(Row(3, Str("A3", "twoCell:", Default)));
        rows.Append(Row(11, Str("A11", "Rotated 20°:", Default)));
        rows.Append(Row(16, Str("A16", "Page break below row 20:", Default)));
        rows.Append(Row(26, Str("A26", "Right: mirrored horizontally (flipH), then vertically (flipV)", Default)));
        rows.Append(Row(27, Str("A27", "End of the sheet", Default)));
        rows.Append(Row(29, Str("A29", "BMP, GIF:", Default)));
        rows.Append(Row(33, Str("A33", "RLE BMP, LZW TIFF:", Default)));
        var cols = "<cols><col min=\"1\" max=\"12\" width=\"13\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(),
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" +
            "<headerFooter differentFirst=\"1\"><oddHeader>&amp;L&amp;G&amp;C&amp;A</oddHeader><firstHeader>&amp;L&amp;G&amp;C&amp;A</firstHeader><firstFooter>&amp;CPage &amp;P of &amp;N&amp;R&amp;G</firstFooter><oddFooter>&amp;CPage &amp;P of &amp;N&amp;R&amp;G</oddFooter></headerFooter>" +
            "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"20\" max=\"16383\" man=\"1\"/></rowBreaks><drawing r:id=\"rId1\"/><legacyDrawingHF r:id=\"rId2\"/>");
    }

    // ------------------------------------------------------------------ conditional formatting

    /// <summary>An icon set that only exists in the Excel 2010 extension list, with percent thresholds written as <c>xm:f</c> formulas.</summary>
    private static string X14IconRule(string sqref, string set, int count, int priority) =>
        "<x14:conditionalFormatting><x14:cfRule type=\"iconSet\" priority=\"" + priority + "\" id=\"{AAAAAAAA-0000-0000-0000-0000000001" + priority + "}\"><x14:iconSet iconSet=\"" + set + "\">" +
        string.Concat(Enumerable.Range(0, count).Select(k => $"<x14:cfvo type=\"percent\"><xm:f>{k * 100 / count}</xm:f></x14:cfvo>")) +
        $"</x14:iconSet></x14:cfRule><xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting>";

    private static string ConditionalSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Conditional formatting", Bold)));
        string[] headers = ["Scores", "Zebra", "Top 3", "Above average", "Duplicates", "Color scale", "Data bar", "Solid bar", "Arrows", "Lights", "Symbols"];
        var header = new StringBuilder();
        for (int c = 0; c < headers.Length; c++) header.Append(Str($"{ColumnName(c)}3", headers[c], Header));
        rows.Append(Row(3, header.ToString()));

        double[] scores = [5, 18, 32, 47, 51, 63, 72, 88, 94, 100];
        double[] mixed = [12, 7, 33, 18, 25, 3, 40, 21, 9, 15];
        double[] signed = [-40, -10, 0, 15, 30, 60, -25, 80, 45, 100];
        string[] colors = ["red", "blue", "red", "green", "blue", "red", "black", "green", "gold", "teal"];
        for (int i = 0; i < 10; i++)
        {
            int r = 4 + i;
            rows.Append(Row(r,
                Num($"A{r}", scores[i]) + Str($"B{r}", $"Stripe {i + 1}") + Num($"C{r}", mixed[i]) + Num($"D{r}", mixed[i]) + Str($"E{r}", colors[i]) +
                Num($"F{r}", scores[i]) + Num($"G{r}", scores[i]) + Num($"H{r}", signed[i]) + Num($"I{r}", scores[i]) + Num($"J{r}", scores[i]) + Num($"K{r}", scores[i]) + Num($"L{r}", scores[i])));
        }
        rows.Append(Row(15, Str("A15", "Priority, stopIfTrue and relative formulas decide which format wins; fills, fonts, borders and number formats come from the rules.")));

        // Time periods are relative to the day the file is converted, so the dates are generated from today.
        double today = (DateTime.Today - new DateTime(1899, 12, 30)).TotalDays;
        rows.Append(Row(17, Str("A17", "Today", Header) + Str("B17", "Yesterday", Header) + Str("C17", "Tomorrow", Header) + Str("D17", "3 weeks ago", Header) + Str("E17", "Next month", Header)));
        rows.Append(Row(18, Num("A18", today, IsoDate) + Num("B18", today - 1, IsoDate) + Num("C18", today + 1, IsoDate) + Num("D18", today - 21, IsoDate) + Num("E18", today + 40, IsoDate)));
        rows.Append(Row(19, Str("A19", "Time-period rules: only today, yesterday and tomorrow are highlighted.")));

        // The other icon sets, one column each, over a ladder of values.
        string[] sets = ["3ArrowsGray", "3Flags", "3TrafficLights2", "3Signs", "3Symbols2", "4Arrows", "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows", "5Rating", "5Quarters"];
        var setHeader = new StringBuilder();
        for (int c = 0; c < sets.Length; c++) setHeader.Append(Str($"{ColumnName(c)}21", sets[c], Header));
        rows.Append(Row(20, Str("A20", "More icon sets", Bold)));
        rows.Append(Row(21, setHeader.ToString()));
        for (int i = 0; i < 6; i++)
        {
            var line = new StringBuilder();
            for (int c = 0; c < sets.Length; c++) line.Append(Num($"{ColumnName(c)}{22 + i}", i * 20));
            rows.Append(Row(22 + i, line.ToString()));
        }

        // Rules that exist only in the Excel 2010 extension list: the 3Stars / 3Triangles / 5Boxes icon sets and a rule with an inline dxf.
        rows.Append(Row(29, Str("A29", "Excel 2010 extension rules", Bold)));
        rows.Append(Row(30, Str("A30", "3Stars", Header) + Str("B30", "3Triangles", Header) + Str("C30", "5Boxes", Header) + Str("D30", "x14 cellIs > 50", Header)));
        for (int i = 0; i < 5; i++)
            rows.Append(Row(31 + i, Num($"A{31 + i}", i * 20) + Num($"B{31 + i}", i * 20) + Num($"C{31 + i}", i * 20) + Num($"D{31 + i}", i * 20)));

        const string Percents = "<cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/><cfvo type=\"percent\" val=\"67\"/>";
        const string X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        var tail = new StringBuilder();
        tail.Append("<conditionalFormatting sqref=\"A4:A13\">" +
                    "<cfRule type=\"cellIs\" dxfId=\"1\" priority=\"1\" operator=\"lessThan\"><formula>20</formula></cfRule>" +
                    "<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"2\" operator=\"greaterThan\"><formula>70</formula></cfRule></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"B4:B13\"><cfRule type=\"expression\" dxfId=\"2\" priority=\"3\"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"C4:C13\"><cfRule type=\"top10\" dxfId=\"3\" priority=\"4\" rank=\"3\"/></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"D4:D13\"><cfRule type=\"aboveAverage\" dxfId=\"1\" priority=\"5\"/></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"E4:E13\"><cfRule type=\"duplicateValues\" dxfId=\"4\" priority=\"6\"/></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"F4:F13\"><cfRule type=\"colorScale\" priority=\"7\"><colorScale><cfvo type=\"min\"/><cfvo type=\"percentile\" val=\"50\"/><cfvo type=\"max\"/>" +
                    "<color rgb=\"FFF8696B\"/><color rgb=\"FFFFEB84\"/><color rgb=\"FF63BE7B\"/></colorScale></cfRule></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"G4:G13\"><cfRule type=\"dataBar\" priority=\"8\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar></cfRule></conditionalFormatting>");
        tail.Append($"<conditionalFormatting sqref=\"H4:H13\"><cfRule type=\"dataBar\" priority=\"9\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF5B9BD5\"/></dataBar>" +
                    $"<extLst><ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"{X14}\"><x14:id>{{AAAAAAAA-0000-0000-0000-000000000001}}</x14:id></ext></extLst></cfRule></conditionalFormatting>");
        tail.Append($"<conditionalFormatting sqref=\"I4:I13\"><cfRule type=\"iconSet\" priority=\"10\"><iconSet iconSet=\"3Arrows\">{Percents}</iconSet></cfRule></conditionalFormatting>");
        tail.Append($"<conditionalFormatting sqref=\"J4:J13\"><cfRule type=\"iconSet\" priority=\"11\"><iconSet iconSet=\"3TrafficLights1\">{Percents}</iconSet></cfRule></conditionalFormatting>");
        tail.Append($"<conditionalFormatting sqref=\"K4:K13\"><cfRule type=\"iconSet\" priority=\"12\"><iconSet iconSet=\"3Symbols\">{Percents}</iconSet></cfRule></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"A18:E18\"><cfRule type=\"timePeriod\" timePeriod=\"today\" dxfId=\"5\" priority=\"13\"/>" +
                    "<cfRule type=\"timePeriod\" timePeriod=\"yesterday\" dxfId=\"1\" priority=\"14\"/><cfRule type=\"timePeriod\" timePeriod=\"tomorrow\" dxfId=\"0\" priority=\"15\"/></conditionalFormatting>");
        for (int c = 0; c < sets.Length; c++)
        {
            int count = sets[c][0] - '0';
            var thresholds = string.Concat(Enumerable.Range(0, count).Select(k => $"<cfvo type=\"percent\" val=\"{k * 100 / count}\"/>"));
            tail.Append($"<conditionalFormatting sqref=\"{ColumnName(c)}22:{ColumnName(c)}27\"><cfRule type=\"iconSet\" priority=\"{20 + c}\"><iconSet iconSet=\"{sets[c]}\">{thresholds}</iconSet></cfRule></conditionalFormatting>");
        }
        tail.Append("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter);
        tail.Append($"<extLst><ext uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\" xmlns:x14=\"{X14}\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:conditionalFormattings>" +
                    "<x14:conditionalFormatting><x14:cfRule type=\"dataBar\" id=\"{AAAAAAAA-0000-0000-0000-000000000001}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\" border=\"1\" negativeBarBorderColorSameAsPositive=\"0\">" +
                    "<x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:negativeFillColor rgb=\"FFC00000\"/><x14:borderColor rgb=\"FF2F5597\"/><x14:negativeBorderColor rgb=\"FF7F0000\"/></x14:dataBar></x14:cfRule><xm:sqref>H4:H13</xm:sqref></x14:conditionalFormatting>" +
                    X14IconRule("A31:A35", "3Stars", 3, 30) + X14IconRule("B31:B35", "3Triangles", 3, 31) + X14IconRule("C31:C35", "5Boxes", 5, 32) +
                    "<x14:conditionalFormatting><x14:cfRule type=\"cellIs\" priority=\"33\" operator=\"greaterThan\" id=\"{AAAAAAAA-0000-0000-0000-000000000005}\"><xm:f>50</xm:f>" +
                    "<x14:dxf><font><b/><color rgb=\"FF7030A0\"/></font><fill><patternFill><bgColor rgb=\"FFE4DFEC\"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>D31:D35</xm:sqref></x14:conditionalFormatting>" +
                    "</x14:conditionalFormattings></ext></extLst>");
        var cols = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"12\" width=\"13\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), tail.ToString());
    }

    private static string PivotSheet(string title)
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", title, Bold)));
        rows.Append(Row(3, Str("A3", "Row Labels", Bold) + Str("B3", "Sum of Sales", Bold)));
        (string Region, double Sales)[] data = [("North", 1250), ("South", 980.5), ("East", 1610), ("West", 745)];
        for (int i = 0; i < data.Length; i++) rows.Append(Row(4 + i, Str($"A{4 + i}", data[i].Region) + Num($"B{4 + i}", data[i].Sales, Money)));
        rows.Append(Row(8, Str("A8", "Grand Total", Bold) + Num("B8", data.Sum(d => d.Sales), Money)));
        return Sheet("<cols><col min=\"1\" max=\"2\" width=\"20\" customWidth=\"1\"/></cols>", rows.ToString(), "<pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    private static void WritePivot(ZipArchive zip, int sheetNumber, string style, string formats = "", bool stripes = false)
    {
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}pivotTable\" Target=\"../pivotTables/pivotTable{sheetNumber}.xml\"/></Relationships>");
        Entry(zip, $"xl/pivotTables/pivotTable{sheetNumber}.xml",
            $"<pivotTableDefinition xmlns=\"{Main}\" name=\"PivotTable{sheetNumber}\" cacheId=\"1\"><location ref=\"A3:B8\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>" +
            "<rowFields count=\"1\"><field x=\"0\"/></rowFields><rowItems count=\"5\"><i><x/></i><i><x v=\"1\"/></i><i><x v=\"2\"/></i><i><x v=\"3\"/></i><i t=\"grand\"><x/></i></rowItems>" + formats +
            (style.Length == 0 ? "" : $"<pivotTableStyleInfo name=\"{style}\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"{(stripes ? 1 : 0)}\" showColStripes=\"0\"/>") + "</pivotTableDefinition>");
    }

    /// <summary>A pivot table with a page field above it and two data fields side by side; formats reach the page item and one data field of one region.</summary>
    private static string PivotFieldsSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(2, Str("A2", "Year", Bold) + Num("B2", 2024)));
        rows.Append(Row(4, Str("A4", "Row Labels", Bold) + Str("B4", "Sum of Sales", Bold) + Str("C4", "Sum of Units", Bold)));
        (string Region, double Sales, double Units)[] data = [("North", 1250, 31), ("South", 980.5, 24), ("East", 1610, 40), ("West", 745, 18)];
        for (int i = 0; i < data.Length; i++) rows.Append(Row(5 + i, Str($"A{5 + i}", data[i].Region) + Num($"B{5 + i}", data[i].Sales, Money) + Num($"C{5 + i}", data[i].Units)));
        rows.Append(Row(9, Str("A9", "Grand Total", Bold) + Num("B9", data.Sum(d => d.Sales), Money) + Num("C9", data.Sum(d => d.Units))));
        rows.Append(Row(11, Str("A11", "Pivot table with a page field (its item in green) and two data fields (the units of East in pink)", Bold)));
        return Sheet("<cols><col min=\"1\" max=\"3\" width=\"20\" customWidth=\"1\"/></cols>", rows.ToString(), "<pageSetup paperSize=\"9\"/>" + HeaderFooter);
    }

    private static void WritePivotFields(ZipArchive zip, int sheetNumber)
    {
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}pivotTable\" Target=\"../pivotTables/pivotTable{sheetNumber}.xml\"/></Relationships>");
        Entry(zip, $"xl/pivotTables/pivotTable{sheetNumber}.xml",
            $"<pivotTableDefinition xmlns=\"{Main}\" name=\"PivotTable{sheetNumber}\" cacheId=\"1\"><location ref=\"A4:C9\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\" rowPageCount=\"1\" colPageCount=\"1\"/>" +
            "<rowFields count=\"1\"><field x=\"0\"/></rowFields><rowItems count=\"5\"><i><x/></i><i><x v=\"1\"/></i><i><x v=\"2\"/></i><i><x v=\"3\"/></i><i t=\"grand\"><x/></i></rowItems>" +
            "<colFields count=\"1\"><field x=\"-2\"/></colFields><colItems count=\"2\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i></colItems>" +
            "<pageFields count=\"1\"><pageField fld=\"3\" item=\"0\" hier=\"-1\"/></pageFields><dataFields count=\"2\"><dataField name=\"Sum of Sales\" fld=\"1\"/><dataField name=\"Sum of Units\" fld=\"2\"/></dataFields>" +
            "<formats count=\"3\">" +
            "<format dxfId=\"1\"><pivotArea dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"3\" count=\"1\" selected=\"0\"><x v=\"0\"/></reference></references></pivotArea></format>" +
            "<format dxfId=\"0\"><pivotArea field=\"0\" type=\"normal\"><references count=\"2\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"1\"/></reference><reference field=\"0\" count=\"1\" selected=\"0\"><x v=\"2\"/></reference></references></pivotArea></format>" +
            "<format dxfId=\"1\"><pivotArea field=\"3\" type=\"button\" dataOnly=\"0\" labelOnly=\"1\" outline=\"0\" axis=\"axisPage\" fieldPosition=\"0\"/></format></formats></pivotTableDefinition>");
    }

    private static string Limits() =>
        Sheet("", Row(1, Num("A1", 50) + Num("B1", 18) + Num("C1", 63)) + Row(2, Num("B2", 100)), "<pageSetup paperSize=\"9\"/>" + HeaderFooter);

    /// <summary>Rules that read other sheets and names, custom icons and data-bar axis options.</summary>
    private static string ConditionalLinksSheet()
    {
        const string X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        const string Percents = "<cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/><cfvo type=\"percent\" val=\"67\"/>";
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Conditional formatting that reads other sheets and names, custom icons and data-bar axes", Bold)));
        rows.Append(Row(3, Str("A3", "> Limits!A1", Header) + Str("B3", "> Limit (name)", Header) + Str("C3", "In list", Header) + Str("D3", "Custom icons", Header) +
                           Str("E3", "Axis middle", Header) + Str("F3", "No axis", Header) + Str("G3", "Right to left", Header)));
        double[] values = [10, 18, 42, 50, 63, 77, 90];
        double[] signed = [-80, -40, -10, 0, 20, 60, 100];
        for (int i = 0; i < values.Length; i++)
        {
            int r = 4 + i;
            rows.Append(Row(r, Num($"A{r}", values[i]) + Num($"B{r}", values[i]) + Num($"C{r}", values[i]) + Num($"D{r}", i * 15) + Num($"E{r}", signed[i]) + Num($"F{r}", signed[i]) + Num($"G{r}", values[i])));
        }
        rows.Append(Row(12, Str("A12", "Limits sheet (hidden): A1 = 50 (the name Limit), B1:C1 = 18, 63 (the list).")));

        var tail = new StringBuilder();
        tail.Append("<conditionalFormatting sqref=\"A4:A10\"><cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>A4&gt;Limits!$A$1</formula></cfRule></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"B4:B10\"><cfRule type=\"expression\" dxfId=\"1\" priority=\"2\"><formula>B4&gt;Limit</formula></cfRule></conditionalFormatting>");
        tail.Append("<conditionalFormatting sqref=\"C4:C10\"><cfRule type=\"expression\" dxfId=\"3\" priority=\"3\"><formula>COUNTIF(Limits!$B$1:$C$1,C4)&gt;0</formula></cfRule></conditionalFormatting>");
        string Linked(string id, string type, string body) =>
            $"<cfRule type=\"{type}\" priority=\"{id.Length}\">{body}<extLst><ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"{X14}\"><x14:id>{{{id}}}</x14:id></ext></extLst></cfRule>";
        tail.Append("<conditionalFormatting sqref=\"D4:D10\">" + Linked("BBBBBBBB-0000-0000-0000-000000000001", "iconSet", $"<iconSet iconSet=\"3Arrows\">{Percents}</iconSet>") + "</conditionalFormatting>");
        int priority = 10;
        string Bar(string column, string id) =>
            $"<conditionalFormatting sqref=\"{column}4:{column}10\"><cfRule type=\"dataBar\" priority=\"{++priority}\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF5B9BD5\"/></dataBar>" +
            $"<extLst><ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"{X14}\"><x14:id>{{{id}}}</x14:id></ext></extLst></cfRule></conditionalFormatting>";
        tail.Append(Bar("E", "BBBBBBBB-0000-0000-0000-000000000002") + Bar("F", "BBBBBBBB-0000-0000-0000-000000000003") + Bar("G", "BBBBBBBB-0000-0000-0000-000000000004"));
        string DataBarExt(string id, string sqref, string attributes, string children) =>
            $"<x14:conditionalFormatting><x14:cfRule type=\"dataBar\" id=\"{{{id}}}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\" {attributes}><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/>{children}</x14:dataBar></x14:cfRule><xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting>";
        tail.Append("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter);
        tail.Append($"<extLst><ext uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\" xmlns:x14=\"{X14}\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:conditionalFormattings>" +
                    // custom icons: a flag for the low level, no icon in the middle, a star for the top
                    "<x14:conditionalFormatting><x14:cfRule type=\"iconSet\" id=\"{BBBBBBBB-0000-0000-0000-000000000001}\"><x14:iconSet iconSet=\"3Arrows\" custom=\"1\">" +
                    "<x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"percent\"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type=\"percent\"><xm:f>67</xm:f></x14:cfvo>" +
                    "<x14:cfIcon iconSet=\"3Flags\" iconId=\"0\"/><x14:cfIcon iconSet=\"NoIcons\" iconId=\"0\"/><x14:cfIcon iconSet=\"3Stars\" iconId=\"2\"/></x14:iconSet></x14:cfRule><xm:sqref>D4:D10</xm:sqref></x14:conditionalFormatting>" +
                    DataBarExt("BBBBBBBB-0000-0000-0000-000000000002", "E4:E10", "axisPosition=\"middle\"", "<x14:negativeFillColor rgb=\"FFC00000\"/><x14:axisColor rgb=\"FF7F0000\"/>") +
                    DataBarExt("BBBBBBBB-0000-0000-0000-000000000003", "F4:F10", "axisPosition=\"none\"", "<x14:negativeFillColor rgb=\"FFC00000\"/>") +
                    DataBarExt("BBBBBBBB-0000-0000-0000-000000000004", "G4:G10", "direction=\"rightToLeft\"", "") +
                    "</x14:conditionalFormattings></ext></extLst>");
        var cols = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"7\" width=\"16\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), tail.ToString());
    }

    /// <summary>Rules that use lookup and reference functions, whole-column references and a data bar that exists only as an Excel 2010 extension.</summary>
    private static string ConditionalTablesSheet()
    {
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Conditional formatting with structured table references and a 3-D reference", Bold) + Num("H1", 10)));
        rows.Append(Row(3, Str("A3", "Item", Header) + Str("B3", "Qty", Header) + Str("C3", "Limit", Header) + Str("D3", "Stock", Header) + Str("E3", "Reorder", Header)));
        string[] items = ["Bolts", "Nuts", "Screws", "Washers"];
        double[] reorder = [11, 2, 9, 16], qty = [12, 3, 9, 20], limit = [10, 5, 9, 15], stock = [4, 8, 40, 60];
        for (int i = 0; i < items.Length; i++)
        {
            int r = 4 + i;
            rows.Append(Row(r, Str($"A{r}", items[i]) + Num($"B{r}", qty[i]) + Num($"C{r}", limit[i]) + Num($"D{r}", stock[i]) + Num($"E{r}", reorder[i])));
        }
        rows.Append(Row(8, Str("A8", "Total") + Num("B8", 44) + Num("C8", 39) + Num("D8", 112) + Num("E8", 38)));
        rows.Append(Row(10, Str("A10", "Qty above the Limit of its row: Stock[@Limit]")));
        rows.Append(Row(11, Str("A11", "Limit above the table's average: AVERAGE(Stock[Limit])")));
        rows.Append(Row(12, Str("A12", "Stock below the three H1 cells: SUM('Conditional tables:Span B'!$H$1)")));
        rows.Append(Row(13, Str("A13", "Reorder above the Limit of its row, with no table name: [@Limit] (mint)")));
        string Rule(string sqref, int priority, string formula) =>
            $"<conditionalFormatting sqref=\"{sqref}\"><cfRule type=\"expression\" dxfId=\"7\" priority=\"{priority}\"><formula>{SecurityElement.Escape(formula)}</formula></cfRule></conditionalFormatting>";
        var tail = Rule("B4:B7", 1, "B4>Stock[@Limit]") + Rule("C4:C7", 2, "C4>AVERAGE(Stock[Limit])") + Rule("D4:D7", 3, "D4<SUM('Conditional tables:Span B'!$H$1)")
            + "<conditionalFormatting sqref=\"E4:E7\"><cfRule type=\"expression\" dxfId=\"8\" priority=\"4\"><formula>E4&gt;[@Limit]</formula></cfRule></conditionalFormatting>"
            + "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"landscape\"/>" + HeaderFooter;
        return Sheet("<cols><col min=\"1\" max=\"5\" width=\"24\" customWidth=\"1\"/></cols>", rows.ToString(), tail);
    }

    private static void WriteConditionalTables(ZipArchive zip, int sheetNumber)
    {
        Entry(zip, "xl/tables/table1.xml",
            $"<table xmlns=\"{Main}\" id=\"1\" name=\"Stock\" displayName=\"Stock\" ref=\"A3:E8\" totalsRowCount=\"1\"><tableColumns count=\"5\">" +
            "<tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Qty\"/><tableColumn id=\"3\" name=\"Limit\"/><tableColumn id=\"4\" name=\"Stock\"/><tableColumn id=\"5\" name=\"Reorder\"/></tableColumns></table>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}table\" Target=\"../tables/table1.xml\"/></Relationships>");
    }

    private static string ConditionalLookupsSheet()
    {
        const string X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        string[] names = ["Ann", "Bob", "Cy", "Ann", "Dee", "Bob", "Eli"];
        double[] scores = [72, 35, 58, 40, 91, 66, 24];
        var rows = new StringBuilder();
        rows.Append(Row(1, Str("A1", "Conditional formatting with lookups, whole columns and an x14-only data bar", Bold)));
        rows.Append(Row(3, Str("A3", "Name: repeats", Header) + Str("B3", "Score: x14 bar", Header) + Str("C3", "VLOOKUP > 60", Header) + Str("D3", "Name total > 100", Header) +
                           Str("E3", "INDEX/MATCH", Header) + Str("F3", "OFFSET > 60", Header) + Str("G3", "INDIRECT < 30", Header) +
                           Str("H3", "SUMIF > 100", Header) + Str("I3", "COUNTIFS = 1", Header) + Str("J3", "TEXT = 35", Header) + Str("K3", "INDIRECT R1C1 < 30", Header) + Str("L3", "R1C1 column < 30", Header)));
        for (int i = 0; i < names.Length; i++)
        {
            int r = 4 + i;
            rows.Append(Row(r, Str($"A{r}", names[i]) + Num($"B{r}", scores[i]) + Num($"C{r}", scores[i]) + Num($"D{r}", scores[i]) + Num($"E{r}", scores[i]) + Num($"F{r}", scores[i]) + Num($"G{r}", scores[i]) +
                           Num($"H{r}", scores[i]) + Num($"I{r}", scores[i]) + Num($"J{r}", scores[i]) + Num($"K{r}", scores[i])));
        }
        rows.Append(Row(12, Str("A12", "Repeated names (whole-column COUNTIF), scores of repeated names add up, Dee's score is looked up by MATCH.")));
        string Rule(string sqref, int dxf, int priority, string formula) =>
            $"<conditionalFormatting sqref=\"{sqref}\"><cfRule type=\"expression\" dxfId=\"{dxf}\" priority=\"{priority}\"><formula>{SecurityElement.Escape(formula)}</formula></cfRule></conditionalFormatting>";
        var tail = new StringBuilder();
        tail.Append(Rule("A4:A10", 4, 1, "COUNTIF($A:$A,A4)>1"));
        tail.Append(Rule("C4:C10", 1, 2, "VLOOKUP($A4,$A$4:$B$10,2,FALSE)>60"));
        tail.Append(Rule("D4:D10", 3, 3, "SUMPRODUCT(($A$4:$A$10=$A4)*($B$4:$B$10))>100"));
        tail.Append(Rule("E4:E10", 1, 4, "INDEX($B$4:$B$10,MATCH(\"Dee\",$A$4:$A$10,0))=B4"));
        tail.Append(Rule("F4:F10", 0, 5, "OFFSET(F4,0,-4)>60"));
        tail.Append(Rule("G4:G10", 5, 6, "INDIRECT(\"B\"&ROW())<30"));
        // The functions of milestone 6, all in the lavender format: Ann and Bob (the repeated names) total over 100 only for Ann (112) and Bob (101).
        tail.Append(Rule("H4:H10", 6, 8, "SUMIF($A$4:$A$10,$A4,$B$4:$B$10)>100"));
        tail.Append(Rule("I4:I10", 6, 9, "COUNTIFS($A$4:$A$10,$A4,$B$4:$B$10,\">50\")=1"));
        tail.Append(Rule("J4:J10", 6, 10, "TEXT(J4,\"0\")=\"35\""));
        tail.Append(Rule("K4:K10", 6, 11, "INDIRECT(\"RC2\",FALSE)<30"));
        // A whole column in R1C1 text (C[-1], the column to the left), indexed by row: only Eli (24) is under 30.
        tail.Append(Rule("L4:L10", 6, 12, "INDEX(INDIRECT(\"C[-1]\",FALSE),ROW())<30"));
        tail.Append("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>" + HeaderFooter);
        tail.Append($"<extLst><ext uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\" xmlns:x14=\"{X14}\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:conditionalFormattings>" +
                    "<x14:conditionalFormatting><x14:cfRule type=\"dataBar\" priority=\"7\" id=\"{CCCCCCCC-0000-0000-0000-000000000001}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\" border=\"1\">" +
                    "<x14:cfvo type=\"num\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"num\"><xm:f>100</xm:f></x14:cfvo><x14:borderColor rgb=\"FF2E75B6\"/></x14:dataBar></x14:cfRule><xm:sqref>B4:B10</xm:sqref></x14:conditionalFormatting>" +
                    "</x14:conditionalFormattings></ext></extLst>");
        var cols = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><cols><col min=\"1\" max=\"12\" width=\"16\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(), tail.ToString());
    }

    private static string Hidden() =>
        Sheet("", Row(1, Str("A1", HiddenMarker)), "<pageSetup paperSize=\"9\"/>" + HeaderFooter);

    private static void Entry(ZipArchive zip, string name, string content)
    {
        using var s = zip.CreateEntry(name).Open();
        s.Write(Encoding.UTF8.GetBytes(content));
    }
}
