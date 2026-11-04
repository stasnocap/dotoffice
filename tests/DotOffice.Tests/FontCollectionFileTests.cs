using System.Buffers.Binary;
using DotOffice.Fonts;
using Xunit;

namespace DotOffice.Tests;

public class FontCollectionFileTests
{
    /// <summary>Wraps a TrueType font into a one-face .ttc (table offsets in a collection are file-absolute).</summary>
    private static byte[] MakeTtc(byte[] ttf)
    {
        const int header = 16;
        var data = new byte[header + ttf.Length];
        ttf.CopyTo(data, header);
        BinaryPrimitives.WriteUInt32BigEndian(data, 0x74746366);
        BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(4), 0x00010000);
        BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(8), 1);
        BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(12), header);
        int n = BinaryPrimitives.ReadUInt16BigEndian(ttf.AsSpan(4));
        for (int i = 0; i < n; i++)
        {
            var off = data.AsSpan(header + 12 + i * 16 + 8);
            BinaryPrimitives.WriteUInt32BigEndian(off, BinaryPrimitives.ReadUInt32BigEndian(off) + header);
        }
        return data;
    }

    [Fact]
    public void LoadsFaceFromCollection()
    {
        var ttf = File.ReadAllBytes(TestAssets.Path("NotoSans-Regular.ttf"));
        var plain = TrueTypeFont.Load(ttf);
        var face = TrueTypeFont.Load(MakeTtc(ttf));
        Assert.Equal(plain.NumGlyphs, face.NumGlyphs);
        Assert.Equal(plain.GetGlyphId('A'), face.GetGlyphId('A'));
        Assert.Throws<InvalidDataException>(() => TrueTypeFont.Load(MakeTtc(ttf), 1));
    }

    [Fact]
    public void IndexesCollectionFiles()
    {
        var dir = Directory.CreateTempSubdirectory().FullName;
        try
        {
            File.WriteAllBytes(Path.Combine(dir, "a.ttc"), MakeTtc(File.ReadAllBytes(TestAssets.Path("NotoSans-Regular.ttf"))));
            var faces = FontFaceInfo.TryReadAll(Path.Combine(dir, "a.ttc"));
            Assert.Single(faces);
            Assert.Equal("Noto Sans", faces[0].FamilyName);
        }
        finally { Directory.Delete(dir, true); }
    }
}
