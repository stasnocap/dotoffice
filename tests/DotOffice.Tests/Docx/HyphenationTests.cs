using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class HyphenationTests
{
    private const string Run = "<w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr>";
    private const string On = "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:autoHyphenation/></w:settings>";

    private static string Para(string text, string pPr = "") => $"<w:p><w:pPr>{pPr}</w:pPr><w:r>{Run}<w:t>{text}</w:t></w:r></w:p>";

    private static string With(string options) => On.Replace("<w:autoHyphenation/>", "<w:autoHyphenation/>" + options);

    private static List<string> Lines(string body, bool hyphenate, string? options = null)
    {
        var doc = DocxReader.Read(new DocxBuilder { Body = body + Section(), Settings = options is not null ? With(options) : hyphenate ? On : null }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
        return Texts(page).GroupBy(t => Math.Round(t.Baseline, 1)).OrderBy(g => g.Key)
            .Select(g => string.Concat(g.OrderBy(t => t.X).Select(t => t.Text)).TrimEnd()).ToList();
    }

    private const string Word = "characteristically";

    // Fillers of growing length until the long word straddles the end of the first line.
    private static string Filled(int words) => string.Join(' ', Enumerable.Repeat("lorem", words)) + " " + Word;

    private static int Straddling() =>
        Enumerable.Range(1, 14).First(n => Lines(Para(Filled(n)), hyphenate: false) is [_, Word, ..]);

    [Fact]
    public void Off_by_default_the_word_moves_to_the_next_line()
    {
        var lines = Lines(Para(Filled(Straddling())), hyphenate: false);
        Assert.Equal(Word, lines[1]);
        Assert.DoesNotContain(lines, l => l.EndsWith('-'));
    }

    [Fact]
    public void Long_word_is_split_at_a_pattern_point_with_a_hyphen()
    {
        var lines = Lines(Para(Filled(Straddling())), hyphenate: true);
        Assert.EndsWith("-", lines[0]);
        string head = lines[0].Split(' ')[^1].TrimEnd('-');
        Assert.Equal(Word, head + lines[1]);
        Assert.Contains(head.Length, DotOffice.Layout.Hyphenation.Hyphenator.ForLanguage("en-US")!.Find(Word));
    }

    [Fact]
    public void Suppress_auto_hyphens_keeps_the_word_whole()
    {
        var lines = Lines(Para(Filled(Straddling()), "<w:suppressAutoHyphens/>"), hyphenate: true);
        Assert.Equal(Word, lines[1]);
    }

    [Fact]
    public void Patterns_keep_two_letters_before_and_three_after()
    {
        var h = DotOffice.Layout.Hyphenation.Hyphenator.ForLanguage("en-US")!;
        Assert.Equal([4, 6, 9], h.Find(Word).Take(3));
        Assert.Empty(h.Find("pine"));
    }

    [Fact]
    public void Zone_wider_than_the_line_keeps_every_word_whole()
    {
        var lines = Lines(Para(Filled(Straddling())), true, "<w:hyphenationZone w:val=\"20000\"/>");
        Assert.Equal(Word, lines[1]);
    }

    [Fact]
    public void Zero_zone_hyphenates_wherever_the_word_straddles()
    {
        var lines = Lines(Para(Filled(Straddling())), true, "<w:hyphenationZone w:val=\"0\"/>");
        Assert.EndsWith("-", lines[0]);
    }

    [Fact]
    public void Capitals_are_kept_whole_with_do_not_hyphenate_caps()
    {
        string caps = Filled(Straddling()).Replace(Word, Word.ToUpperInvariant());
        Assert.EndsWith("-", Lines(Para(caps), true, "<w:hyphenationZone w:val=\"0\"/>")[0]);
        Assert.Equal(Word.ToUpperInvariant(), Lines(Para(caps), true, "<w:hyphenationZone w:val=\"0\"/><w:doNotHyphenateCaps/>")[1]);
    }

    [Fact]
    public void Consecutive_hyphen_limit_stops_a_run_of_hyphenated_lines()
    {
        string text = string.Join(' ', Enumerable.Repeat("internationalization characteristically", 8));
        int Longest(int indent, string options)
        {
            int best = 0, cur = 0;
            foreach (var l in Lines(Para(text, $"<w:ind w:right=\"{indent}\"/>"), true, "<w:hyphenationZone w:val=\"0\"/>" + options))
            {
                cur = l.EndsWith('-') ? cur + 1 : 0;
                best = Math.Max(best, cur);
            }
            return best;
        }

        // Some line width makes several lines in a row end in a hyphen; the limit then caps that run.
        int indent = Enumerable.Range(50, 80).Select(n => n * 100).First(n => Longest(n, "") >= 2);
        Assert.Equal(1, Longest(indent, "<w:consecutiveHyphenLimit w:val=\"1\"/>"));
    }
}
