using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;
using static DotOffice.Dml.DrawingMl;

namespace DotOffice.Pptx;

/// <summary>
/// Tables (<c>a:tbl</c> in a graphic frame), drawn as one rectangle per cell with its fill and text, then the cell borders as lines. The
/// table style (from <c>tableStyles.xml</c>, or one of the built-in Medium Style 2 looks) gives fills, borders and text weight and color per
/// table part; a cell's own properties outrank it.
/// </summary>
internal sealed partial class PptxContext
{
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private readonly Dictionary<string, XElement> _tableStyles = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The built-in Medium Style 2 table styles by id: the plain one in dark 1 and the six accent ones.</summary>
    private static readonly Dictionary<string, string> MediumStyle2 = new(StringComparer.OrdinalIgnoreCase)
    {
        ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = "dk1",
        ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = "accent1",
        ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}"] = "accent2",
        ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = "accent3",
        ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = "accent4",
        ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}"] = "accent5",
        ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = "accent6",
    };

    private static XElement MediumStyle2Of(string color)
    {
        string Line(string w) => $"<a:ln w=\"{w}\" cmpd=\"sng\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln>";
        string Fill(string tint) => $"<a:fill><a:solidFill><a:schemeClr val=\"{color}\">{tint}</a:schemeClr></a:solidFill></a:fill>";
        string Strong(string border) => $"<a:tcTxStyle b=\"on\"><a:fontRef idx=\"minor\"/><a:schemeClr val=\"lt1\"/></a:tcTxStyle><a:tcStyle>{border}{Fill("")}</a:tcStyle>";
        string thin = Line("12700");
        var xml = $"""
            <a:tblStyle xmlns:a="{A}">
              <a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>
                <a:left>{thin}</a:left><a:right>{thin}</a:right><a:top>{thin}</a:top><a:bottom>{thin}</a:bottom><a:insideH>{thin}</a:insideH><a:insideV>{thin}</a:insideV>
              </a:tcBdr>{Fill("<a:tint val=\"20000\"/>")}</a:tcStyle></a:wholeTbl>
              <a:band1H><a:tcStyle>{Fill("<a:tint val=\"40000\"/>")}</a:tcStyle></a:band1H>
              <a:band1V><a:tcStyle>{Fill("<a:tint val=\"40000\"/>")}</a:tcStyle></a:band1V>
              <a:lastCol>{Strong("")}</a:lastCol>
              <a:firstCol>{Strong("")}</a:firstCol>
              <a:lastRow>{Strong($"<a:tcBdr><a:top>{Line("38100")}</a:top></a:tcBdr>")}</a:lastRow>
              <a:firstRow>{Strong($"<a:tcBdr><a:bottom>{Line("38100")}</a:bottom></a:tcBdr>")}</a:firstRow>
            </a:tblStyle>
            """;
        return XElement.Parse(xml);
    }

    private XElement? TableStyle(XElement table)
    {
        var tblPr = Child(table, "tblPr");
        // A table style written into the table itself (rare) comes first; then the presentation's table styles, then the built-in ones.
        if (Child(tblPr, "tableStyle") is { } inline) return inline;
        var id = Child(tblPr, "tableStyleId")?.Value.Trim();
        if (string.IsNullOrEmpty(id)) return null;
        if (_tableStyles.TryGetValue(id, out var style) && style.HasElements) return style;
        if (MediumStyle2.TryGetValue(id, out var color)) return MediumStyle2Of(color);
        Note("tables in a built-in table style other than Medium Style 2 are drawn without the style");
        return null;
    }

    private sealed record TableCell(XElement Tc, int Row, int Column, int RowSpan, int ColumnSpan);

    private void ReadTable(Slide slide, XElement table, Scope scope, Placement place, ref int z)
    {
        XNamespace a = A, p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var columns = Child(table, "tblGrid")?.Elements().Where(e => e.Name.LocalName == "gridCol").Select(c => Number(c, "w")).ToList() ?? [];
        var rows = table.Elements().Where(e => e.Name.LocalName == "tr").ToList();
        if (columns.Count == 0 || rows.Count == 0) return;
        var heights = rows.Select(r => Number(r, "h")).ToList();
        double[] xs = new double[columns.Count + 1], ys = new double[rows.Count + 1];
        for (int c = 0; c < columns.Count; c++) xs[c + 1] = xs[c] + columns[c];
        for (int r = 0; r < rows.Count; r++) ys[r + 1] = ys[r] + heights[r];
        double frameCx = place.Width * EmuPerPoint, frameCy = place.Height * EmuPerPoint;
        if (frameCx <= 0) frameCx = xs[^1];
        if (frameCy <= 0) frameCy = ys[^1];

        // The cells that start a merge or stand alone; cells covered by a merge (hMerge, vMerge) are drawn by the cell that starts it.
        var cells = new List<TableCell>();
        for (int r = 0; r < rows.Count; r++)
        {
            var tcs = rows[r].Elements().Where(e => e.Name.LocalName == "tc").ToList();
            for (int c = 0; c < tcs.Count && c < columns.Count; c++)
            {
                var tc = tcs[c];
                if (Flag(tc, "hMerge") || Flag(tc, "vMerge")) continue;
                int colSpan = Math.Clamp((int)Number(tc, "gridSpan", 1), 1, columns.Count - c);
                int rowSpan = Math.Clamp((int)Number(tc, "rowSpan", 1), 1, rows.Count - r);
                cells.Add(new TableCell(tc, r, c, rowSpan, colSpan));
            }
        }

        var tblPr = Child(table, "tblPr");
        bool firstRow = Flag(tblPr, "firstRow"), lastRow = Flag(tblPr, "lastRow"), firstCol = Flag(tblPr, "firstCol"), lastCol = Flag(tblPr, "lastCol");
        bool bandRow = Flag(tblPr, "bandRow"), bandCol = Flag(tblPr, "bandCol");
        var style = TableStyle(table);
        int lastR = rows.Count - 1, lastC = columns.Count - 1;

        // The table parts that apply to a cell, lowest first, each with the cells it covers (top, bottom, left, right).
        List<(XElement Part, int R0, int R1, int C0, int C1)> Parts(TableCell cell)
        {
            var parts = new List<(XElement, int, int, int, int)>();
            if (style is null) return parts;
            void Add(string name, int r0, int r1, int c0, int c1)
            {
                if (Child(style, name) is { } part) parts.Add((part, r0, r1, c0, c1));
            }
            int r = cell.Row, c = cell.Column;
            bool inFirstRow = firstRow && r == 0, inLastRow = lastRow && r + cell.RowSpan - 1 == lastR;
            bool inFirstCol = firstCol && c == 0, inLastCol = lastCol && c + cell.ColumnSpan - 1 == lastC;
            Add("wholeTbl", 0, lastR, 0, lastC);
            if (bandCol && !inFirstCol && !inLastCol) Add((c - (firstCol ? 1 : 0)) % 2 == 0 ? "band1V" : "band2V", 0, lastR, c, c);
            if (bandRow && !inFirstRow && !inLastRow) Add((r - (firstRow ? 1 : 0)) % 2 == 0 ? "band1H" : "band2H", r, r, 0, lastC);
            if (inLastCol) Add("lastCol", 0, lastR, lastC, lastC);
            if (inFirstCol) Add("firstCol", 0, lastR, 0, 0);
            if (inLastRow) Add("lastRow", lastR, lastR, 0, lastC);
            if (inFirstRow) Add("firstRow", 0, 0, 0, lastC);
            if (inFirstRow && inLastCol) Add("neCell", 0, 0, lastC, lastC);
            if (inFirstRow && inFirstCol) Add("nwCell", 0, 0, 0, 0);
            if (inLastRow && inLastCol) Add("seCell", lastR, lastR, lastC, lastC);
            if (inLastRow && inFirstCol) Add("swCell", lastR, lastR, 0, 0);
            return parts;
        }

        Placement Within(double x, double y, double w, double h) =>
            place with { Frame = place.Frame.Within(new FrameFraction(x / frameCx, y / frameCy, w / frameCx, h / frameCy)) };
        string Num(double v) => Math.Round(v).ToString(CultureInfo.InvariantCulture);
        XElement Xfrm(double x, double y, double w, double h) =>
            new(a + "xfrm", new XElement(a + "off", new XAttribute("x", Num(x)), new XAttribute("y", Num(y))), new XElement(a + "ext", new XAttribute("cx", Num(w)), new XAttribute("cy", Num(h))));

        var borders = new List<(XElement Line, double X, double Y, double W, double H)>();
        foreach (var cell in cells)
        {
            double x = xs[cell.Column], y = ys[cell.Row], w = xs[cell.Column + cell.ColumnSpan] - x, h = ys[cell.Row + cell.RowSpan] - y;
            var tcPr = Child(cell.Tc, "tcPr");
            var parts = Parts(cell);

            // Fill: the highest table part that states one, then the cell's own.
            XElement? fill = null;
            foreach (var (part, _, _, _, _) in parts)
                if (Path(part, "tcStyle", "fill")?.Elements().FirstOrDefault() is { } f) fill = f;
                else if (Path(part, "tcStyle", "fillRef") is { } fillRef && fillRef.Elements().FirstOrDefault() is { } refColor) fill = new XElement(a + "solidFill", new XElement(refColor));
            if (tcPr?.Elements().FirstOrDefault(e => FillKinds.Contains(e.Name.LocalName)) is { } own) fill = own;

            // Text: the table parts' text style as the lowest list style, under the cell's own.
            var lstStyle = new XElement(a + "lstStyle");
            var defRPr = new XElement(a + "defRPr");
            foreach (var (part, _, _, _, _) in parts)
            {
                if (Child(part, "tcTxStyle") is not { } text) continue;
                foreach (var name in new[] { "b", "i" })
                    if ((string?)text.Attribute(name) is { } v) defRPr.SetAttributeValue(name, v == "on" ? "1" : "0");
                if (text.Elements().FirstOrDefault(e => e.Name.LocalName is "schemeClr" or "srgbClr" or "sysClr" or "prstClr" or "hslClr" or "scrgbClr") is { } color)
                {
                    foreach (var old in defRPr.Elements().Where(e => FillKinds.Contains(e.Name.LocalName)).ToList()) old.Remove();
                    defRPr.Add(new XElement(a + "solidFill", new XElement(color)));
                }
                if ((string?)Child(text, "fontRef")?.Attribute("idx") == "major")
                {
                    Child(defRPr, "latin")?.Remove();
                    defRPr.Add(new XElement(a + "latin", new XAttribute("typeface", "+mj-lt")));
                }
            }
            var body = Child(cell.Tc, "txBody") is { } tx ? new XElement(tx) : new XElement(a + "txBody", new XElement(a + "p"));
            var ownStyle = Child(body, "lstStyle");
            for (int level = 1; level <= 9; level++)
            {
                var levelStyle = new XElement(a + $"lvl{level}pPr", new XElement(defRPr));
                if (Child(ownStyle, $"lvl{level}pPr") is { } ownLevel)
                {
                    MergeParagraph(levelStyle, ownLevel);
                    MergeRun(Child(levelStyle, "defRPr")!, Child(ownLevel, "defRPr"));
                }
                lstStyle.Add(levelStyle);
            }
            ownStyle?.Remove();
            Child(body, "bodyPr")?.Remove();
            string Margin(string name, double fallback) => Num(tcPr?.Attribute(name) is null ? fallback : Number(tcPr, name));
            body.AddFirst(
                new XElement(a + "bodyPr",
                    new XAttribute("wrap", "square"),
                    new XAttribute("lIns", Margin("marL", 91_440)), new XAttribute("rIns", Margin("marR", 91_440)),
                    new XAttribute("tIns", Margin("marT", 45_720)), new XAttribute("bIns", Margin("marB", 45_720)),
                    new XAttribute("anchor", (string?)tcPr?.Attribute("anchor") ?? "t")),
                lstStyle);
            body.Name = p + "txBody";

            var sp = new XElement(p + "sp",
                new XElement(p + "nvSpPr", new XElement(p + "cNvPr", new XAttribute("id", 0), new XAttribute("name", "")), new XElement(p + "cNvSpPr"), new XElement(p + "nvPr")),
                new XElement(p + "spPr", Xfrm(x, y, w, h), new XElement(a + "prstGeom", new XAttribute("prst", "rect")),
                    fill is null ? new XElement(a + "noFill") : new XElement(fill), new XElement(a + "ln", new XElement(a + "noFill"))),
                body);
            ReadObject(slide, sp, scope with { IsSlide = false }, Within(x, y, w, h), ref z);

            // Borders: each edge takes the outer border of the part when the cell is on that part's edge, its inside border otherwise; then the cell's own line.
            int r0 = cell.Row, r1 = cell.Row + cell.RowSpan - 1, c0 = cell.Column, c1 = cell.Column + cell.ColumnSpan - 1;
            XElement? Edge(string outer, string inside, Func<(XElement Part, int R0, int R1, int C0, int C1), bool> isOuter, string ownName)
            {
                XElement? line = null;
                foreach (var part in parts)
                {
                    var side = Path(part.Part, "tcStyle", "tcBdr", isOuter(part) ? outer : inside);
                    if (Child(side, "ln") is { } ln) line = ln;
                    else if (Child(side, "lnRef") is { } lnRef && lnRef.Elements().FirstOrDefault() is { } color)
                        line = new XElement(a + "ln", new XAttribute("w", "12700"), new XElement(a + "solidFill", new XElement(color)));
                }
                if (Child(tcPr, ownName) is { } own) line = own;
                return line;
            }
            void Border(XElement? line, double bx, double by, double bw, double bh)
            {
                if (line is null || line.Elements().All(e => e.Name.LocalName != "solidFill" && e.Name.LocalName != "gradFill")) return;
                borders.Add((new XElement(line) { Name = a + "ln" }, bx, by, bw, bh));
            }
            Border(Edge("top", "insideH", t => r0 == t.R0, "lnT"), x, y, w, 0);
            Border(Edge("bottom", "insideH", t => r1 == t.R1, "lnB"), x, y + h, w, 0);
            Border(Edge("left", "insideV", t => c0 == t.C0, "lnL"), x, y, 0, h);
            Border(Edge("right", "insideV", t => c1 == t.C1, "lnR"), x + w, y, 0, h);
        }

        // Borders go over every cell fill, so a neighbour's fill does not hide half of a shared edge.
        foreach (var (line, x, y, w, h) in borders)
        {
            var cxn = new XElement(p + "cxnSp",
                new XElement(p + "nvCxnSpPr", new XElement(p + "cNvPr", new XAttribute("id", 0), new XAttribute("name", "")), new XElement(p + "cNvCxnSpPr"), new XElement(p + "nvPr")),
                new XElement(p + "spPr", Xfrm(x, y, w, h), new XElement(a + "prstGeom", new XAttribute("prst", "line")), line));
            ReadObject(slide, cxn, scope with { IsSlide = false }, Within(x, y, w, h), ref z);
        }
    }
}
