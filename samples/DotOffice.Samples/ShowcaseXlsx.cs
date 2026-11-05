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
        Fraction1 = 27, Fraction2 = 28, Eighths = 29, IndexedRed = 30;

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
            ("Conditional", ConditionalSheet(), "visible"),
            ("Paper", PaperSheet(), "visible"),
            ("Areas", AreasSheet(), "visible"),
            ("Comments", CommentsSheet(), "visible"),
            ("Secret", Hidden(), "hidden"),
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
        WriteComments(zip, sheets.FindIndex(x => x.Name == "Comments") + 1, content);
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
            "<definedNames><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"1\">'Long table'!$1:$2</definedName>" +
            $"<definedName name=\"_xlnm.Print_Area\" localSheetId=\"{sheets.FindIndex(x => x.Name == "Areas")}\">Areas!$A$1:$C$3,Areas!$E$5:$F$7,Areas!$A$10:$B$11</definedName></definedNames></workbook>");
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
        "<cellXfs count=\"31\">" +
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
        "</cellXfs>" +
        "<dxfs count=\"6\">" +
        "<dxf><font><color rgb=\"FF9C0006\"/></font><fill><patternFill><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>" +                       // 0 bad
        "<dxf><font><color rgb=\"FF006100\"/></font><fill><patternFill><bgColor rgb=\"FFC6EFCE\"/></patternFill></fill></dxf>" +                       // 1 good
        "<dxf><fill><patternFill><bgColor rgb=\"FFDDEBF7\"/></patternFill></fill></dxf>" +                                                              // 2 zebra
        "<dxf><font><b/><color rgb=\"FF9C5700\"/></font><fill><patternFill><bgColor rgb=\"FFFFEB9C\"/></patternFill></fill>" +
        "<border><bottom style=\"thin\"><color rgb=\"FF9C5700\"/></bottom></border></dxf>" +                                                          // 3 top
        "<dxf><font><b/><i/><u/><color rgb=\"FFC00000\"/></font><numFmt numFmtId=\"167\" formatCode=\"0.0&quot; (!)&quot;\"/></dxf>" +                  // 4 duplicate
        "<dxf><font><b/></font><fill><patternFill><bgColor rgb=\"FFFFD966\"/></patternFill></fill></dxf>" +                                                  // 5 time period
        "</dxfs>" +
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

    private static string Pic(string name, string extra = "", int rotation = 0, string flip = "") =>
        $"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"{name}\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId1\"/>{extra}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>" +
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
            TwoCell(1, 17, 4, 24, Pic("Across the page break")) +                                                                  // continues on the next page
            "<xdr:oneCellAnchor>" + Marker("from", 6, 17) + "<xdr:ext cx=\"1270000\" cy=\"635000\"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"9\" name=\"Shape\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>";
        Entry(zip, $"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{Rel}\">{anchors}</xdr:wsDr>");
        Entry(zip, $"xl/drawings/_rels/drawing{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}image\" Target=\"../media/image1.png\"/></Relationships>");
        Entry(zip, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
        using (var media = zip.CreateEntry("xl/media/image1.png").Open()) media.Write(ShowcaseDocx.GradientPng(160, 100));
        content.Append("<Default Extension=\"png\" ContentType=\"image/png\"/>");
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
        var cols = "<cols><col min=\"1\" max=\"12\" width=\"13\" customWidth=\"1\"/></cols>";
        return Sheet(cols, rows.ToString(),
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\"/>" + HeaderFooter +
            "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"20\" max=\"16383\" man=\"1\"/></rowBreaks><drawing r:id=\"rId1\"/>");
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
                Num($"F{r}", scores[i]) + Num($"G{r}", scores[i]) + Num($"H{r}", signed[i]) + Num($"I{r}", scores[i]) + Num($"J{r}", scores[i]) + Num($"K{r}", scores[i])));
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

    private static string Hidden() =>
        Sheet("", Row(1, Str("A1", HiddenMarker)), "<pageSetup paperSize=\"9\"/>" + HeaderFooter);

    private static void Entry(ZipArchive zip, string name, string content)
    {
        using var s = zip.CreateEntry(name).Open();
        s.Write(Encoding.UTF8.GetBytes(content));
    }
}
