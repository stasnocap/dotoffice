using DotOffice.Layout.Bidi;

namespace DotOffice.Tests.Layout;

/// <summary>
/// Cases from (or in the style of) BidiCharacterTest.txt. The implementation was also checked against the
/// complete Unicode 16 BidiCharacterTest.txt (levels and visual order); that file is not part of the repository.
/// </summary>
public class BidiTests
{
    private static (byte[] Levels, string Visual) Resolve(string text, byte paragraphLevel)
    {
        var cps = text.EnumerateRunes().Select(r => r.Value).ToArray();
        var levels = BidiAlgorithm.ResolveLevels(cps.Select(BidiAlgorithm.GetClass).ToArray(), cps, paragraphLevel);
        var order = BidiAlgorithm.VisualOrder(levels);
        // Characters at odd levels are displayed mirrored (L4).
        return (levels, string.Concat(order.Select(i => char.ConvertFromUtf32(levels[i] % 2 == 1 ? BidiAlgorithm.Mirror(cps[i]) : cps[i]))));
    }

    [Fact]
    public void Classes_and_mirroring_come_from_the_unicode_data()
    {
        Assert.Equal(BidiClass.L, BidiAlgorithm.GetClass('a'));
        Assert.Equal(BidiClass.R, BidiAlgorithm.GetClass('א'));
        Assert.Equal(BidiClass.AL, BidiAlgorithm.GetClass('ب'));
        Assert.Equal(BidiClass.EN, BidiAlgorithm.GetClass('7'));
        Assert.Equal(BidiClass.AN, BidiAlgorithm.GetClass('٣'));
        Assert.Equal(BidiClass.NSM, BidiAlgorithm.GetClass(0x064E));
        Assert.Equal(BidiClass.R, BidiAlgorithm.GetClass(0x05FF)); // unassigned Hebrew block defaults to R
        Assert.Equal(')', BidiAlgorithm.Mirror('('));
        Assert.Equal('«', BidiAlgorithm.Mirror('»'));
        Assert.Equal('a', BidiAlgorithm.Mirror('a'));
    }

    [Fact]
    public void Rtl_run_inside_ltr_paragraph_is_reversed()
    {
        var (levels, visual) = Resolve("ab אבג cd", 0);
        Assert.Equal([0, 0, 0, 1, 1, 1, 0, 0, 0], levels);
        Assert.Equal("ab גבא cd", visual);
    }

    [Fact]
    public void Numbers_keep_their_order_inside_rtl_text()
    {
        var (levels, visual) = Resolve("אב 123 גד", 1);
        Assert.Equal([1, 1, 1, 2, 2, 2, 1, 1, 1], levels);
        Assert.Equal("דג 123 בא", visual);
        // European digits after Arabic letters become Arabic numbers (W2) but still read left to right.
        Assert.Equal("12 ب", Resolve("ب 12", 1).Visual);
    }

    [Fact]
    public void Brackets_take_the_direction_of_their_content_and_context()
    {
        // N0: brackets around English in RTL text take the RTL context and are mirrored, so they still enclose it.
        Assert.Equal("ג (ab) בא", Resolve("אב (ab) ג", 1).Visual);
        Assert.Equal([0, 0, 0, 0, 1, 0, 0, 0], Resolve("ab (א) c", 0).Levels);
        Assert.Equal([1, 1, 1, 1, 1, 1], Resolve("א(ב)ג ", 1).Levels);
    }

    [Fact]
    public void Explicit_embeddings_and_isolates_raise_levels()
    {
        const char rle = '‫', pdf = '‬', rli = '⁧', pdi = '⁩';
        // Latin letters in an RTL embedding are at level 2; the trailing PDF returns to the paragraph level (L1).
        Assert.Equal([0, 0, 0, 2, 2, 0], Resolve($"a {rle}bc{pdf}", 0).Levels);
        var (levels, _) = Resolve($"a{rli}b א{pdi}c", 0);
        Assert.Equal([0, 0, 2, 1, 1, 0, 0], levels);
        // Trailing whitespace returns to the paragraph level (L1).
        Assert.Equal([1, 1, 0, 0], Resolve("אב  ", 0).Levels);
    }
}
