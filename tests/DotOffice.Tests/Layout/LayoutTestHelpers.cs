using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

internal static class LayoutTestHelpers
{
    public static readonly TrueTypeFont Noto = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    public static readonly RunProperties Text = new() { FontFamily = "Noto Sans", FontSize = 10 };

    public static LayoutDocument Render(Document doc, ConversionWarningList? warnings = null) =>
        new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)), Warnings = warnings }).Layout(doc);

    public static Document Doc(params Block[] blocks) => Doc(new Section(), blocks);

    public static Document Doc(Section section, params Block[] blocks)
    {
        var doc = new Document();
        section.Blocks.AddRange(blocks);
        doc.Sections.Add(section);
        return doc;
    }

    public static Paragraph Para(string text, ParagraphProperties? props = null, RunProperties? run = null)
    {
        var p = new Paragraph { Properties = props ?? ParagraphProperties.Default, MarkProperties = run ?? Text };
        if (text.Length > 0) p.Inlines.Add(new TextRun(text, run ?? Text));
        return p;
    }

    public static Paragraph Para(ParagraphProperties props, params Inline[] inlines)
    {
        var p = new Paragraph { Properties = props, MarkProperties = Text };
        p.Inlines.AddRange(inlines);
        return p;
    }

    public static IEnumerable<TextItem> Texts(LayoutPage page) => page.Items.OfType<TextItem>();

    public static string PageText(LayoutPage page) => string.Concat(Texts(page).Select(t => t.Text));

    /// <summary>Visible words grouped by baseline, top to bottom.</summary>
    public static List<(double Baseline, List<TextItem> Items)> Lines(LayoutPage page) =>
        Texts(page).GroupBy(t => Math.Round(t.Baseline, 2)).OrderBy(g => g.Key)
                   .Select(g => (g.Key, g.OrderBy(t => t.X).ToList())).ToList();

    public static double Right(TextItem t) => t.X + t.Font.MeasureText(t.Text, t.FontSize) + t.CharacterSpacing * t.Text.Length;

    public static string Words(int count, string prefix = "word") =>
        string.Join(' ', Enumerable.Range(1, count).Select(i => $"{prefix}{i}"));
}