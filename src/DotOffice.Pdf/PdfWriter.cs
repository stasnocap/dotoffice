using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace DotOffice.Pdf;

/// <summary>Low-level PDF serializer: numbered indirect objects, streams, xref table and trailer.</summary>
internal sealed class PdfWriter
{
    private readonly Stream _output;
    private readonly List<long> _offsets = [0];
    private long _position;

    public PdfWriter(Stream output)
    {
        _output = output;
        Write("%PDF-1.7\n");
        WriteBytes([(byte)'%', 0xE2, 0xE3, 0xCF, 0xD3, (byte)'\n']);
    }

    /// <summary>Reserves an object number so it can be referenced before it is written.</summary>
    public int Allocate()
    {
        _offsets.Add(-1);
        return _offsets.Count - 1;
    }

    public void WriteObject(int id, string body)
    {
        _offsets[id] = _position;
        Write($"{id} 0 obj\n{body}\nendobj\n");
    }

    public void WriteStream(int id, string dictionaryEntries, ReadOnlySpan<byte> data, bool compress = true)
    {
        byte[]? compressed = null;
        if (compress)
        {
            using var buffer = new MemoryStream();
            using (var zlib = new ZLibStream(buffer, CompressionLevel.Optimal, leaveOpen: true))
                zlib.Write(data);
            compressed = buffer.ToArray();
        }

        var payload = compressed is null ? data : compressed;
        var filter = compressed is null ? "" : " /Filter /FlateDecode";
        _offsets[id] = _position;
        Write($"{id} 0 obj\n<< /Length {payload.Length}{filter} {dictionaryEntries} >>\nstream\n");
        WriteBytes(payload);
        Write("\nendstream\nendobj\n");
    }

    public void Finish(int catalogId, int infoId)
    {
        long xref = _position;
        var sb = new StringBuilder();
        sb.Append($"xref\n0 {_offsets.Count}\n");
        sb.Append("0000000000 65535 f \n");
        for (int i = 1; i < _offsets.Count; i++)
        {
            if (_offsets[i] < 0) throw new InvalidOperationException($"Object {i} was allocated but never written.");
            sb.Append(_offsets[i].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        sb.Append($"trailer\n<< /Size {_offsets.Count} /Root {catalogId} 0 R /Info {infoId} 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        Write(sb.ToString());
    }

    private void Write(string text) => WriteBytes(Encoding.Latin1.GetBytes(text));

    private void WriteBytes(ReadOnlySpan<byte> bytes)
    {
        _output.Write(bytes);
        _position += bytes.Length;
    }
}

internal static class PdfFormat
{
    public static string Number(double value)
    {
        var rounded = Math.Round(value, 4);
        return rounded == 0 ? "0" : rounded.ToString("0.####", CultureInfo.InvariantCulture);
    }

    /// <summary>Encodes text as a PDF literal string, using UTF-16BE when it is not plain ASCII.</summary>
    public static string TextString(string text)
    {
        if (text.All(c => c >= 32 && c < 127))
            return "(" + text.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)") + ")";
        return "<FEFF" + Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(text)) + ">";
    }
}
