namespace DotOffice.Fonts.Shaping;

/// <summary>
/// The legacy <c>kern</c> table (OpenType flavour, version 0) that HarfBuzz applies to fonts without GPOS
/// (hb-ot-kern-table.hh and hb-kern.hh). Only format 0 (sorted pairs) is read; other formats are ignored.
/// </summary>
internal sealed class KernTable
{
    private readonly record struct SubTable(int Offset, int Count, bool CrossStream);

    private readonly byte[] _data;
    private readonly List<SubTable> _subTables = [];

    public KernTable(ReadOnlySpan<byte> table)
    {
        _data = table.ToArray();
        if (_data.Length < 4 || U16(0) != 0) return;
        int count = U16(2), pos = 4;
        for (int i = 0; i < count && pos + 6 <= _data.Length; i++)
        {
            int length = U16(pos + 2), format = _data[pos + 4], coverage = _data[pos + 5];
            if (length < 6) break;
            // Format 0 is a header (nPairs, searchRange, entrySelector, rangeShift) followed by 6-byte pairs.
            if (format == 0 && (coverage & 1) != 0 && pos + 14 <= _data.Length)
            {
                int pairs = Math.Min(U16(pos + 6), (_data.Length - pos - 14) / 6);
                _subTables.Add(new SubTable(pos + 14, pairs, (coverage & 4) != 0));
            }
            pos += length;
        }
    }

    public bool IsEmpty => _subTables.Count == 0;

    /// <summary>True when a subtable moves the second glyph vertically instead of changing advances.</summary>
    public bool HasCrossStream => _subTables.Any(s => s.CrossStream);

    /// <summary>The sum of the kerning of the glyph pair over all subtables, split by cross-stream flag.</summary>
    public (int Horizontal, int CrossStream) Kerning(int left, int right)
    {
        int horizontal = 0, cross = 0;
        uint key = (uint)(left << 16 | right);
        foreach (var (offset, count, crossStream) in _subTables)
        {
            int lo = 0, hi = count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2, at = offset + mid * 6;
                uint pair = (uint)(U16(at) << 16 | U16(at + 2));
                if (pair < key) lo = mid + 1;
                else if (pair > key) hi = mid - 1;
                else
                {
                    int value = (short)U16(at + 4);
                    if (crossStream) cross += value; else horizontal += value;
                    break;
                }
            }
        }
        return (horizontal, cross);
    }

    private int U16(int at) => _data[at] << 8 | _data[at + 1];
}
