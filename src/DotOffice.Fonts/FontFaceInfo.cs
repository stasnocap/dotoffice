using System.Buffers.Binary;
using System.Text;

namespace DotOffice.Fonts;

/// <summary>
/// Lightweight description of a font file read from its name/head/OS2 tables only, so that
/// whole font directories can be indexed without parsing glyph data.
/// </summary>
public sealed record FontFaceInfo(string Path, string FamilyName, string TypographicFamilyName, string SubfamilyName,
    int Weight, int WidthClass, bool IsItalic)
{
    public bool IsBold => Weight >= 600;

    /// <summary>Reads face info from a TrueType (glyf) font file; returns null for anything unsupported.</summary>
    public static FontFaceInfo? TryRead(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096);
            return TryRead(stream, path);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidDataException
                                      or ArgumentOutOfRangeException or EndOfStreamException)
        {
            return null;
        }
    }

    private static FontFaceInfo? TryRead(Stream stream, string path)
    {
        Span<byte> header = stackalloc byte[12];
        stream.ReadExactly(header);
        uint version = BinaryPrimitives.ReadUInt32BigEndian(header);
        if (version != 0x00010000 && version != 0x74727565) return null; // CFF and collections unsupported
        int numTables = BinaryPrimitives.ReadUInt16BigEndian(header[4..]);
        if (numTables is 0 or > 512) return null;

        var directory = new byte[numTables * 16];
        stream.ReadExactly(directory);
        byte[]? name = null, head = null, os2 = null;
        bool hasGlyf = false;
        for (int i = 0; i < numTables; i++)
        {
            var rec = directory.AsSpan(i * 16, 16);
            var tag = Encoding.ASCII.GetString(rec[..4]);
            uint offset = BinaryPrimitives.ReadUInt32BigEndian(rec[8..]);
            uint length = BinaryPrimitives.ReadUInt32BigEndian(rec[12..]);
            if (tag == "glyf") hasGlyf = true;
            if (tag is not ("name" or "head" or "OS/2") || length > 1 << 20 || offset + (long)length > stream.Length)
                continue;
            var data = new byte[length];
            stream.Position = offset;
            stream.ReadExactly(data);
            switch (tag)
            {
                case "name": name = data; break;
                case "head": head = data; break;
                default: os2 = data; break;
            }
        }
        if (!hasGlyf || name is null || head is null || head.Length < 54) return null;

        var names = FontNames.Parse(name);
        if (names.GetValueOrDefault(1) is not { } family) return null;
        int macStyle = BinaryPrimitives.ReadUInt16BigEndian(head.AsSpan(44));
        int weight = (macStyle & 1) != 0 ? 700 : 400, width = 5;
        bool italic = (macStyle & 2) != 0;
        if (os2 is { Length: >= 64 })
        {
            weight = BinaryPrimitives.ReadUInt16BigEndian(os2.AsSpan(4));
            width = BinaryPrimitives.ReadUInt16BigEndian(os2.AsSpan(6));
            int fsSelection = BinaryPrimitives.ReadUInt16BigEndian(os2.AsSpan(62));
            italic |= (fsSelection & 1) != 0;
            if ((fsSelection & 0x20) != 0 && weight < 600) weight = 700;
        }
        return new FontFaceInfo(path, family, names.GetValueOrDefault(16) ?? family,
            names.GetValueOrDefault(2) ?? "Regular", weight, width, italic);
    }
}

internal static class FontNames
{
    /// <summary>Parses family (1), subfamily (2), PostScript (6) and typographic family/subfamily (16/17) names.</summary>
    public static Dictionary<int, string> Parse(ReadOnlySpan<byte> table)
    {
        var result = new Dictionary<int, string>();
        if (table.Length < 6) return result;
        int count = BinaryPrimitives.ReadUInt16BigEndian(table[2..]);
        int strings = BinaryPrimitives.ReadUInt16BigEndian(table[4..]);
        var english = new HashSet<int>();
        for (int i = 0; i < count && 6 + i * 12 + 12 <= table.Length; i++)
        {
            var rec = table.Slice(6 + i * 12, 12);
            int platform = BinaryPrimitives.ReadUInt16BigEndian(rec);
            int encoding = BinaryPrimitives.ReadUInt16BigEndian(rec[2..]);
            int language = BinaryPrimitives.ReadUInt16BigEndian(rec[4..]);
            int nameId = BinaryPrimitives.ReadUInt16BigEndian(rec[6..]);
            int length = BinaryPrimitives.ReadUInt16BigEndian(rec[8..]);
            int offset = BinaryPrimitives.ReadUInt16BigEndian(rec[10..]);
            if (nameId is not (1 or 2 or 6 or 16 or 17) || strings + offset + length > table.Length) continue;
            var bytes = table.Slice(strings + offset, length);
            string? value = (platform, encoding) switch
            {
                (3, 1) or (3, 10) or (0, _) => Encoding.BigEndianUnicode.GetString(bytes),
                (1, 0) => Encoding.Latin1.GetString(bytes),
                _ => null,
            };
            if (string.IsNullOrWhiteSpace(value)) continue;
            // Prefer Windows US-English names, but accept anything as a fallback.
            bool isEnglish = platform == 3 && language == 0x409;
            if (!result.ContainsKey(nameId) || (isEnglish && english.Add(nameId)))
                result[nameId] = value.Trim();
        }
        return result;
    }
}