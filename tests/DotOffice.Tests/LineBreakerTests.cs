using DotOffice.Layout.LineBreaking;

namespace DotOffice.Tests;

public class LineBreakerTests
{
    /// <summary>Splits text at every allowed break, marking mandatory breaks with '!'.</summary>
    private static string Segments(string text)
    {
        var parts = new List<string>();
        int last = 0;
        foreach (var (index, mandatory) in LineBreaker.GetBreaks(text))
        {
            parts.Add(text[last..index] + (mandatory ? "!" : ""));
            last = index;
        }
        parts.Add(text[last..]);
        return string.Join("|", parts);
    }

    [Theory]
    [InlineData("Hello world", "Hello |world")]
    [InlineData("Hello   world", "Hello   |world")]
    [InlineData("well-known", "well-|known")]
    [InlineData("pages 10-20", "pages |10-20")]          // HY × NU
    [InlineData("end. Next", "end. |Next")]                // no break before '.'
    [InlineData("(see) this", "(see) |this")]              // OP ×, × CP
    [InlineData("a ( b", "a |( b")]                        // OP SP* ×
    [InlineData("10 kg", "10 kg")]               // NBSP glues
    [InlineData("$100 is 50%", "$100 |is |50%")]           // PR × NU, NU × PO
    [InlineData("3.14159", "3.14159")]
    [InlineData("a\tb", "a\t|b")]                          // break after tab, not before
    [InlineData("x​z", "x​|z")]                  // zero width space
    [InlineData("“Quoted” text", "“Quoted” |text")]
    public void Breaks_latin_text(string text, string expected) => Assert.Equal(expected, Segments(text));

    [Fact]
    public void Breaks_between_ideographs_but_not_before_closing_punctuation()
    {
        Assert.Equal("中|文|字", Segments("中文字"));
        Assert.Equal("中|文。|字", Segments("中文。字"));   // no break before 。(CL)
        Assert.Equal("「中|文」", Segments("「中文」"));     // no break after 「 (OP)
        Assert.Equal("ひ|ら|が|な", Segments("ひらがな"));
    }

    [Fact]
    public void Line_feeds_are_mandatory_breaks()
    {
        Assert.Equal("a\n!|b", Segments("a\nb"));
        Assert.Equal("a\r\n!|b", Segments("a\r\nb"));     // CR × LF
        Assert.Equal("a\u2028!|b", Segments("a\u2028b"));
        Assert.Equal("a\f!|b", Segments("a\fb"));
    }

    [Fact]
    public void Combining_marks_and_surrogates_stay_attached()
    {
        Assert.Equal("é |x", Segments("é x"));
        Assert.Equal("😀|😀", Segments("😀😀"));             // emoji are ID: breakable, but never inside a pair
        Assert.Equal("👍🏽", Segments("👍🏽"));                 // EB × EM
    }

    [Fact]
    public void Object_replacement_allows_breaks_on_both_sides()
    {
        Assert.Equal("ab|￼|cd", Segments("ab￼cd"));
    }

    [Fact]
    public void Class_table_lookup_matches_unicode_data()
    {
        Assert.Equal(LineBreakClass.AL, LineBreaker.GetClass('A'));
        Assert.Equal(LineBreakClass.SP, LineBreaker.GetClass(' '));
        Assert.Equal(LineBreakClass.GL, LineBreaker.GetClass(0x00A0));
        Assert.Equal(LineBreakClass.ID, LineBreaker.GetClass(0x4E2D));
        Assert.Equal(LineBreakClass.HY, LineBreaker.GetClass('-'));
        Assert.Equal(LineBreakClass.XX, LineBreaker.GetClass(0x10FFFF));
    }
}