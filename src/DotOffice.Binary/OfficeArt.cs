using System.Buffers.Binary;
using System.IO.Compression;
using DotOffice.Core;

namespace DotOffice.Binary;

/// <summary>Office Art (Escher) pictures shared by the binary formats: the BLIP store of a drawing group and the BLIP records in it.</summary>
internal static class OfficeArt
{
    private const int MaxPicture = 64 << 20;

    /// <summary>
    /// The pictures of a DggContainer's BStoreContainer, in order (shapes refer to them by 1-based index); null where a picture is
    /// missing, kept outside the record (a delay stream) or in a format that is not read.
    /// </summary>
    public static List<ImageData?> ReadStore(byte[] escher, byte[]? delay = null)
    {
        var result = new List<ImageData?>();
        foreach (var (type, _, body, length) in Records(escher, 0, escher.Length))
        {
            if (type != 0xF000) continue;
            foreach (var (t, _, b, l) in Records(escher, body, body + length))
            {
                if (t != 0xF001) continue;
                foreach (var (bt, _, bb, bl) in Records(escher, b, b + l))
                {
                    if (bt != 0xF007) continue;
                    result.Add(Bse(escher, bb, bl, delay));
                    if (result.Count > 100_000) return result;
                }
            }
        }
        return result;
    }

    /// <summary>The picture of one FBSE record body, or null when it is missing or not read.</summary>
    public static ImageData? Bse(byte[] escher, int body, int length, byte[]? delay = null)
    {
        // FBSE: 36 bytes, then the name (cbName at offset 33), then the BLIP record itself.
        if (length < 36) return null;
        ImageData? image = null;
        int at = body + 36 + escher[body + 33];
        foreach (var blip in Records(escher, at, body + length)) { image = Decode(escher, blip.Type, blip.Inst, blip.Body, blip.Length); break; }
        // Word keeps the BLIP in its WordDocument stream, at the FBSE's foDelay.
        long foDelay = BinaryPrimitives.ReadUInt32LittleEndian(escher.AsSpan(body + 28));
        if (image is null && at >= body + length && delay is not null && foDelay + 8 <= delay.Length)
            foreach (var blip in Records(delay, (int)foDelay, delay.Length)) { image = Decode(delay, blip.Type, blip.Inst, blip.Body, blip.Length); break; }
        return image;
    }

    /// <summary>Decodes one BLIP record body: JPEG, PNG, DIB (given a bitmap file header), TIFF, or an EMF, WMF or PICT metafile (inflated).</summary>
    public static ImageData? Decode(byte[] escher, int type, int inst, int body, int length)
    {
        try
        {
            var span = escher.AsSpan(body, length);
            // One 16-byte UID, or two when the instance's low bit is set; bitmaps add a tag byte, metafiles a 34-byte header.
            int uids = (inst & 1) != 0 ? 32 : 16;
            switch (type)
            {
                case 0xF01D or 0xF01E or 0xF02A or 0xF029:
                    return span.Length > uids + 1 ? new ImageData(span[(uids + 1)..].ToArray()) : null;
                case 0xF01F:
                    return span.Length > uids + 1 ? Dib(span[(uids + 1)..]) : null;
                case 0xF01A or 0xF01B or 0xF01C:
                    if (span.Length < uids + 34) return null;
                    var header = span[uids..];
                    int compressed = (int)Math.Min(BinaryPrimitives.ReadUInt32LittleEndian(header[28..]), (uint)(span.Length - uids - 34));
                    var data = span.Slice(uids + 34, compressed);
                    var bytes = header[32] == 0 ? Inflate(data, (int)Math.Min(BinaryPrimitives.ReadUInt32LittleEndian(header), MaxPicture)) : data.ToArray();
                    return bytes is null ? null : new ImageData(bytes);
            }
        }
        catch (InvalidDataException) { }
        return null;
    }

    private static byte[]? Inflate(ReadOnlySpan<byte> data, int expected)
    {
        using var input = new MemoryStream(data.ToArray());
        using var z = new ZLibStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        int n;
        while ((n = z.Read(buffer)) > 0)
        {
            output.Write(buffer, 0, n);
            if (output.Length > Math.Max(expected, 1) + (1 << 20) || output.Length > MaxPicture) return null;
        }
        return output.ToArray();
    }

    /// <summary>A packed DIB becomes a BMP file: a 14-byte file header pointing past the info header and the color table.</summary>
    private static ImageData? Dib(ReadOnlySpan<byte> dib)
    {
        if (dib.Length < 40) return null;
        int headerSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(dib);
        int bits = BinaryPrimitives.ReadUInt16LittleEndian(dib[14..]);
        int used = headerSize >= 40 ? (int)BinaryPrimitives.ReadUInt32LittleEndian(dib[32..]) : 0;
        int colors = bits <= 8 ? (used > 0 ? used : 1 << bits) : used;
        int masks = headerSize == 40 && BinaryPrimitives.ReadUInt32LittleEndian(dib[16..]) == 3 ? 12 : 0;
        var file = new byte[14 + dib.Length];
        file[0] = (byte)'B'; file[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(2), (uint)file.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(10), (uint)(14 + headerSize + masks + colors * 4));
        dib.CopyTo(file.AsSpan(14));
        return new ImageData(file);
    }

    /// <summary>The records between <paramref name="start"/> and <paramref name="end"/>; a length past the end is cut to it.</summary>
    public static IEnumerable<(int Type, int Inst, int Body, int Length)> Records(byte[] escher, int start, int end)
    {
        int at = start;
        while (at + 8 <= end)
        {
            int verInst = BinaryPrimitives.ReadUInt16LittleEndian(escher.AsSpan(at)), type = BinaryPrimitives.ReadUInt16LittleEndian(escher.AsSpan(at + 2));
            long length = BinaryPrimitives.ReadUInt32LittleEndian(escher.AsSpan(at + 4));
            int body = at + 8;
            if (length > end - body) length = end - body;
            yield return (type, verInst >> 4, body, (int)length);
            at = body + (int)length;
        }
    }

    public readonly record struct Rect(int Left, int Top, int Right, int Bottom);

    public sealed class EscherShape
    {
        public int Type;
        public bool Group, Patriarch, FlipH, FlipV, HasClientData, HasTextbox;
        public int Spid, ClientAnchorAt = -1, ClientAnchorLength;
        public Rect? ChildRect, GroupRect;
        public readonly Dictionary<int, uint> Properties = [];
    }

    public static FrameFraction Relative(Rect child, Rect space)
    {
        double w = space.Right - space.Left, h = space.Bottom - space.Top;
        if (w <= 0 || h <= 0) return FrameFraction.Whole;
        return new FrameFraction((child.Left - space.Left) / w, (child.Top - space.Top) / h, (child.Right - child.Left) / w, (child.Bottom - child.Top) / h);
    }

    public static EscherShape ReadShape(byte[] escher, int start, int end)
    {
        var shape = new EscherShape();
        Rect ReadRect(int at) => new(BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at)), BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at + 4)),
            BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at + 8)), BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at + 12)));
        foreach (var (type, inst, body, length) in Records(escher, start, end))
        {
            switch (type)
            {
                case 0xF009 when length >= 16: shape.GroupRect = ReadRect(body); break;
                case 0xF00A when length >= 8:
                    uint flags = BinaryPrimitives.ReadUInt32LittleEndian(escher.AsSpan(body + 4));
                    shape.Type = inst;
                    shape.Spid = BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(body));
                    shape.Group = (flags & 1) != 0; shape.Patriarch = (flags & 4) != 0;
                    shape.FlipH = (flags & 0x40) != 0; shape.FlipV = (flags & 0x80) != 0;
                    break;
                case 0xF00B or 0xF122:
                    for (int i = 0; i < inst && body + 6 * i + 6 <= body + length; i++)
                    {
                        int id = BinaryPrimitives.ReadUInt16LittleEndian(escher.AsSpan(body + 6 * i));
                        if ((id & 0x8000) == 0) shape.Properties[id & 0x3FFF] = BinaryPrimitives.ReadUInt32LittleEndian(escher.AsSpan(body + 6 * i + 2));
                    }
                    break;
                case 0xF00F when length >= 16: shape.ChildRect = ReadRect(body); break;
                case 0xF010: (shape.ClientAnchorAt, shape.ClientAnchorLength) = (body, length); break;
                case 0xF011: shape.HasClientData = true; break;
                case 0xF00D: shape.HasTextbox = true; break;
            }
        }
        return shape;
    }

}
