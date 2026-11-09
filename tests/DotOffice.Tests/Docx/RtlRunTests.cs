using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

/// <summary>
/// Runs with w:rtl in paragraphs with and without w:bidi, against Word 2016's output for the same lines
/// (Arial 16 pt, every run w:rtl): each stretch of w:rtl runs is ordered as a right-to-left embedding.
/// </summary>
public class RtlRunTests
{
    private const string H = "שלום";
    private static readonly TrueTypeFont Hebrew = TrueTypeFont.Load(TestAssets.Path("NotoSansHebrew-Regular.ttf"));

    private static string Mirror(string text) => string.Concat(text.Select(c => c switch { '(' => ')', ')' => '(', '[' => ']', ']' => '[', _ => c }));

    private static string Rev(string text) => string.Concat(text.Reverse());

    /// <summary>The glyphs drawn on the line from left to right: right-to-left items are reversed with mirrored brackets.</summary>
    private static string Visual(string text, bool bidi = false)
    {
        var props = bidi ? "<w:pPr><w:bidi/></w:pPr>" : "";
        var run = $"<w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\" w:cs=\"Noto Sans Hebrew\"/><w:rtl/></w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
        var doc = DocxReader.Read(new DocxBuilder { Body = $"<w:p>{props}{run}</w:p>" + Section() }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto).Add(Hebrew)) }).Layout(doc).Pages[0];
        return string.Concat(Texts(page).OrderBy(t => t.X).Select(t => t.RightToLeft ? Rev(Mirror(t.Text)) : t.Text));
    }

    private static readonly string R = Rev(H);

    private static void Both(string expected, string text)
    {
        Assert.Equal(expected, Visual(text));
        Assert.Equal(expected, Visual(text, bidi: true));
    }

    [Fact]
    public void Hebrew_abc_hebrew_keeps_its_spaces_and_each_word_in_place() => Both($"{R} abc {R}", $"{H} abc {H}");

    [Fact]
    public void Hebrew_digits_hebrew() => Both($"{R} 123 {R}", $"{H} 123 {H}");

    // Word: "H (abc) H" with balanced brackets; they take the embedding's direction and are mirrored.
    [Fact]
    public void Hebrew_parenthesized_latin_hebrew_keeps_balanced_brackets() => Both($"{R} (abc) {R}", $"{H} (abc) {H}");

    // Word: ".H - 1,234.50 ,H"
    [Fact]
    public void Punctuation_and_number_between_hebrew_words_run_right_to_left() => Both($".{R} - 1,234.50 ,{R}", $"{H}, 1,234.50 - {H}.");

    // Word: "?! [H] (H)"
    [Fact]
    public void Brackets_and_marks_around_hebrew() => Both($"?! [{R}] ({R})", $"({H}) [{H}] !?");

    // Word: "... 123 H abc" (it also leaves the trailing space, which the layout trims, at the line's left end).
    [Fact]
    public void Latin_then_hebrew_numbers_and_dots_reads_right_to_left() =>
        Both($"... 123 {R} abc", $"abc {H} 123 ... ");

    /// <summary>Like <see cref="Visual"/> for a paragraph of several runs, each with its own w:rtl flag.</summary>
    private static string VisualRuns(bool bidi, params (string Text, bool Rtl)[] runs)
    {
        var props = bidi ? "<w:pPr><w:bidi/></w:pPr>" : "";
        var body = string.Concat(runs.Select(r =>
            $"<w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\" w:cs=\"Noto Sans Hebrew\"/>{(r.Rtl ? "<w:rtl/>" : "")}</w:rPr><w:t xml:space=\"preserve\">{r.Text}</w:t></w:r>"));
        var doc = DocxReader.Read(new DocxBuilder { Body = $"<w:p>{props}{body}</w:p>" + Section() }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto).Add(Hebrew)) }).Layout(doc).Pages[0];
        return string.Concat(Texts(page).OrderBy(t => t.X).Select(t => t.RightToLeft ? Rev(Mirror(t.Text)) : t.Text));
    }

    // The mixed cases below are the paragraphs of Word 2016's rendering (Arial 16 pt) of rtl-run-latin-and-mixed-runs.docx.

    // Word, w:bidi paragraph: w:rtl "H abc " then plain "x y" reads "x y abc H": the plain run is not joined to the Latin of the w:rtl run.
    [Fact]
    public void Plain_latin_after_rtl_run_in_bidi_paragraph_stands_left_of_it() =>
        Assert.Equal($"x y abc {R}", VisualRuns(true, ($"{H} abc ", true), ("x y", false)));

    // Word, w:bidi paragraph: plain "abc ", w:rtl "H xyz", plain " def" reads "def xyz H abc"
    // with no space between def and xyz or between H and abc (the plain runs' spaces lie at the outer ends).
    [Fact]
    public void Rtl_run_between_plain_latin_runs_in_bidi_paragraph() =>
        Assert.Equal($" defxyz {R}abc ", VisualRuns(true, ("abc ", false), ($"{H} xyz", true), (" def", false)));

    // Word, left-to-right paragraph: plain "abc ", w:rtl "H xyz", plain " def" reads "abc xyz H def".
    [Fact]
    public void Rtl_run_between_plain_latin_runs_in_ltr_paragraph() =>
        Assert.Equal($"abc xyz {R} def", VisualRuns(false, ("abc ", false), ($"{H} xyz", true), (" def", false)));

    // Word, left-to-right paragraph: w:rtl "H abc " then plain "x y" reads "abc H" (the space left of abc) followed by "x y".
    [Fact]
    public void Plain_latin_after_rtl_run_in_ltr_paragraph_follows_it() =>
        Assert.Equal($" abc {R}x y", VisualRuns(false, ($"{H} abc ", true), ("x y", false)));

    // Word, left-to-right paragraph: w:rtl "abc ", plain H, w:rtl " def" reads "abc H def": each w:rtl run is reversed on its own
    // (its space goes to its left for "abc ", to its right for " def") and the runs keep their order.
    [Fact]
    public void Rtl_runs_around_a_plain_run_are_not_reversed_together_with_it() =>
        Assert.Equal($" abc{R}def ", VisualRuns(false, ("abc ", true), (H, false), (" def", true)));

    // Word, w:bidi paragraph: w:rtl "abc ", plain H, w:rtl " def" reads "def H abc" (the whole paragraph runs right to left).
    [Fact]
    public void Rtl_runs_around_a_plain_run_in_bidi_paragraph() =>
        Assert.Equal($"def {R} abc", VisualRuns(true, ("abc ", true), (H, false), (" def", true)));
}
