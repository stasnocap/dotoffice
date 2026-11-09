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
}
