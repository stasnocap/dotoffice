using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

public class KernTableTests
{
    // version 0, one subtable: format 0, horizontal, pairs (1,2) = -80 and (3,4) = 25 (cross-stream table has (5,6) = 10)
    private static byte[] Table(params (int Left, int Right, int Value, bool Cross)[] subTables)
    {
        var bytes = new List<byte> { 0, 0, 0, (byte)subTables.Length };
        foreach (var (left, right, value, cross) in subTables)
        {
            bytes.AddRange([0, 0, 0, 20, 0, (byte)(cross ? 5 : 1)]); // version, length, format 0, coverage
            bytes.AddRange([0, 1, 0, 0, 0, 0, 0, 0]);               // nPairs and the (unused) search fields
            bytes.AddRange([(byte)(left >> 8), (byte)left, (byte)(right >> 8), (byte)right, (byte)(value >> 8), (byte)value]);
        }
        return [.. bytes];
    }

    [Fact]
    public void Pairs_are_looked_up_per_subtable_and_added_up()
    {
        var kern = new KernTable(Table((1, 2, -80, false), (1, 2, -20, false), (1, 2, 10, true)));
        Assert.Equal((-100, 10), kern.Kerning(1, 2));
        Assert.Equal((0, 0), kern.Kerning(2, 1));
        Assert.True(kern.HasCrossStream);
    }

    [Fact]
    public void Missing_or_unsupported_tables_are_empty()
    {
        Assert.True(new KernTable([]).IsEmpty);
        Assert.True(new KernTable([0, 1, 0, 0]).IsEmpty); // Apple version
    }
}
