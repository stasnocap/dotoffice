using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using DotOffice.Core;
using DotOffice.Xlsx;

namespace DotOffice.Binary;

/// <summary>Reads Excel 97-2003 workbooks (.xls, BIFF8) into the <see cref="Workbook"/> model, using the cached cell values.</summary>
public static class XlsReader
{
    /// <summary>Reads an XLS file. Corrupt, encrypted, unsupported or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Workbook Read(Stream stream, BinaryReaderOptions? options = null)
    {
        options ??= new BinaryReaderOptions();
        var file = new CompoundFile(stream, options);
        if (file.Contains("EncryptedPackage")) throw new InvalidDataException("The workbook is encrypted.");
        var data = file.ReadStream("Workbook");
        if (data is null)
        {
            if (file.Contains("Book")) throw new InvalidDataException("Excel 5.0/95 (BIFF5) workbooks are not supported.");
            throw new InvalidDataException("The file has no workbook stream; it is not an Excel workbook.");
        }
        try
        {
            return new XlsContext(data, options).Read();
        }
        catch (Exception e) when (e is IndexOutOfRangeException or ArgumentException or OverflowException or FormatException)
        {
            throw new InvalidDataException("The workbook is damaged.", e);
        }
    }
}

internal sealed partial class XlsContext(byte[] data, BinaryReaderOptions options)
{
    private const int MaxCellText = 4096, MaxRows = 65536, MaxColumns = 256;

    private readonly record struct Record(int Id, int Offset, int Length);

    private readonly List<Record> _records = [];
    private readonly List<(RunProperties Font, int Index)> _fonts = [];
    private readonly Dictionary<int, string> _formats = [];
    private readonly List<Xf> _xfs = [];
    private readonly Dictionary<int, CellStyle> _styleCache = [];
    private readonly List<string> _sst = [];
    private readonly Color[] _palette = DefaultPalette.ToArray();
    private bool _date1904;
    private long _cells;

    private sealed record Xf(int Font, int Format, int Align, int Rotation, int Indent, int Borders, int BorderColors, int TopBottomColors, int Fill, int FillColors);

    private ReadOnlySpan<byte> Data(Record r) => data.AsSpan(r.Offset, r.Length);

    // ------------------------------------------------------------------ records

    private void Split()
    {
        int at = 0;
        while (at + 4 <= data.Length)
        {
            int id = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(at)), length = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(at + 2));
            if (at + 4 + length > data.Length) break; // a truncated last record is dropped
            _records.Add(new Record(id, at + 4, length));
            at += 4 + length;
            if (_records.Count > 20_000_000) throw new InvalidDataException("The workbook has too many records.");
        }
    }

    public Workbook Read()
    {
        Split();
        if (_records.Count == 0 || _records[0].Id != 0x0809 && _records[0].Id != 0x0009)
            throw new InvalidDataException("The workbook stream does not start with a BOF record.");
        int version = BinaryPrimitives.ReadUInt16LittleEndian(Data(_records[0]));
        if (version != 0x0600) throw new InvalidDataException("Only Excel 97-2003 (BIFF8) workbooks are supported.");

        var sheets = new List<(string Name, int Offset, int State, int Type)>();
        int i = 0, depth = 0;
        for (; i < _records.Count; i++)
        {
            var r = _records[i];
            var d = Data(r);
            if (r.Id is 0x0809 or 0x0009) { depth++; if (depth > 1) { i = SkipSubstream(i); depth--; } continue; }
            if (r.Id == 0x000A) { if (--depth <= 0) break; continue; }
            switch (r.Id)
            {
                case 0x002F: throw new InvalidDataException("The workbook is encrypted.");
                case 0x0022: _date1904 = d.Length >= 2 && BinaryPrimitives.ReadUInt16LittleEndian(d) == 1; break;
                case 0x0031: ReadFont(d); break;
                case 0x0041: ReadFormat(d); break;
                case 0x00E0: ReadXf(d); break;
                case 0x0092: ReadPalette(d); break;
                case 0x00FC: i = ReadSst(i); break;
                case 0x00EB: i = ReadDrawingGroup(i); break;
                case 0x0085:
                    if (d.Length >= 8)
                    {
                        int offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(d), state = d[4] & 3, type = d[5];
                        var (name, _) = ReadShortString(d, 6);
                        sheets.Add((name, offset, state, type));
                    }
                    break;
            }
        }

        var workbook = new Workbook();
        foreach (var (name, offset, state, type) in sheets)
        {
            if (type != 0) continue; // chart sheets, macro sheets and modules have no cells
            int start = _records.FindIndex(r => r.Offset - 4 == offset);
            if (start < 0 || _records[start].Id is not (0x0809 or 0x0009)) continue;
            var sheet = new Worksheet(name)
            {
                Visibility = state switch { 1 => SheetVisibility.Hidden, 2 => SheetVisibility.VeryHidden, _ => SheetVisibility.Visible },
                DefaultColumnWidth = (8 * 7 + 5) * 0.75,
                DefaultRowHeight = 12.75,
            };
            ReadSheet(sheet, start);
            workbook.Sheets.Add(sheet);
        }
        if (workbook.Sheets.Count == 0) throw new InvalidDataException("The workbook has no worksheets.");
        return workbook;
    }

    /// <summary>The index of the EOF that closes the substream opened at <paramref name="bof"/>.</summary>
    private int SkipSubstream(int bof)
    {
        int depth = 0;
        for (int i = bof; i < _records.Count; i++)
        {
            if (_records[i].Id is 0x0809 or 0x0009) depth++;
            else if (_records[i].Id == 0x000A && --depth == 0) return i;
        }
        return _records.Count - 1;
    }

    // ------------------------------------------------------------------ strings

    private static (string Text, int End) ReadShortString(ReadOnlySpan<byte> d, int at)
    {
        int cch = d[at], flags = d[at + 1];
        return ReadChars(d, at + 2, cch, (flags & 1) != 0);
    }

    private static (string Text, int End) ReadString16(ReadOnlySpan<byte> d, int at)
    {
        int cch = BinaryPrimitives.ReadUInt16LittleEndian(d[at..]), flags = d[at + 2];
        int p = at + 3;
        int runs = 0, ext = 0;
        if ((flags & 8) != 0) { runs = BinaryPrimitives.ReadUInt16LittleEndian(d[p..]); p += 2; }
        if ((flags & 4) != 0) { ext = (int)BinaryPrimitives.ReadUInt32LittleEndian(d[p..]); p += 4; }
        var (text, end) = ReadChars(d, p, cch, (flags & 1) != 0);
        return (text, end + runs * 4 + ext);
    }

    private static (string Text, int End) ReadChars(ReadOnlySpan<byte> d, int at, int cch, bool wide)
    {
        int bytes = wide ? cch * 2 : cch;
        if (at + bytes > d.Length) { bytes = d.Length - at; cch = wide ? bytes / 2 : bytes; }
        return (wide ? Encoding.Unicode.GetString(d.Slice(at, cch * 2)) : Encoding.Latin1.GetString(d.Slice(at, cch)), at + bytes);
    }

    /// <summary>The shared string table: the SST record and the CONTINUE records after it, which may split a string anywhere.</summary>
    private int ReadSst(int index)
    {
        var segments = new List<(int Offset, int Length)> { (_records[index].Offset, _records[index].Length) };
        int last = index;
        while (last + 1 < _records.Count && _records[last + 1].Id == 0x003C)
        {
            last++;
            segments.Add((_records[last].Offset, _records[last].Length));
        }
        if (segments[0].Length < 8) return last;
        uint unique = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(segments[0].Offset + 4));
        int seg = 0, pos = 8;
        long total = 0;

        bool Ensure()
        {
            while (pos >= segments[seg].Length)
            {
                if (seg + 1 >= segments.Count) return false;
                pos -= segments[seg].Length;
                seg++;
                pos = 0;
                return true;
            }
            return true;
        }
        int U8() => Ensure() ? data[segments[seg].Offset + pos++] : -1;
        int U16() { int a = U8(), b = U8(); return a < 0 || b < 0 ? -1 : a | (b << 8); }
        long U32() { long a = U16(), b = U16(); return a < 0 || b < 0 ? -1 : a | (b << 16); }
        void Skip(long n)
        {
            while (n > 0)
            {
                if (!Ensure()) return;
                int take = (int)Math.Min(n, segments[seg].Length - pos);
                pos += take;
                n -= take;
            }
        }

        for (uint s = 0; s < unique && s < 4_000_000; s++)
        {
            int cch = U16();
            if (cch < 0) break;
            int flags = U8();
            if (flags < 0) break;
            int runs = 0;
            long ext = 0;
            if ((flags & 8) != 0) runs = U16();
            if ((flags & 4) != 0) ext = U32();
            bool wide = (flags & 1) != 0;
            var sb = new StringBuilder(Math.Min(cch, MaxCellText));
            int remaining = cch;
            while (remaining > 0)
            {
                if (pos >= segments[seg].Length)
                {
                    if (seg + 1 >= segments.Count) break;
                    seg++;
                    pos = 0;
                    if (segments[seg].Length == 0) break;
                    wide = (data[segments[seg].Offset] & 1) != 0; // a continued string repeats its encoding flag
                    pos = 1;
                }
                int available = segments[seg].Length - pos;
                int chars = Math.Min(remaining, wide ? available / 2 : available);
                if (chars <= 0) break;
                if (sb.Length < MaxCellText)
                {
                    var span = data.AsSpan(segments[seg].Offset + pos, wide ? chars * 2 : chars);
                    sb.Append(wide ? Encoding.Unicode.GetString(span) : Encoding.Latin1.GetString(span));
                }
                pos += wide ? chars * 2 : chars;
                remaining -= chars;
            }
            Skip(runs * 4L);
            Skip(ext);
            total += cch;
            if (total > options.MaxStreamSize) throw new InvalidDataException("The shared string table is larger than the limit.");
            _sst.Add(sb.Length > MaxCellText ? sb.ToString(0, MaxCellText) : sb.ToString());
        }
        return last;
    }

    // ------------------------------------------------------------------ styles

    private void ReadFont(ReadOnlySpan<byte> d)
    {
        if (d.Length < 16) { _fonts.Add((new RunProperties { FontFamily = "Arial", FontSize = 10 }, _fonts.Count)); return; }
        int height = BinaryPrimitives.ReadUInt16LittleEndian(d), flags = BinaryPrimitives.ReadUInt16LittleEndian(d[2..]);
        int color = BinaryPrimitives.ReadUInt16LittleEndian(d[4..]), weight = BinaryPrimitives.ReadUInt16LittleEndian(d[6..]);
        int script = BinaryPrimitives.ReadUInt16LittleEndian(d[8..]), underline = d[10];
        var name = d.Length > 15 ? ReadShortString(d, 14).Text : "Arial";
        var font = new RunProperties
        {
            FontFamily = name, FontSize = Math.Max(1, height / 20.0), Bold = weight >= 600, Italic = (flags & 2) != 0, Strike = (flags & 8) != 0,
            Underline = underline switch { 0 => UnderlineStyle.None, 2 or 0x22 => UnderlineStyle.Double, _ => UnderlineStyle.Single },
            VerticalPosition = script switch { 1 => VerticalPosition.Superscript, 2 => VerticalPosition.Subscript, _ => VerticalPosition.Baseline },
            Color = ColorOf(color, null),
        };
        _fonts.Add((font, _fonts.Count));
    }

    private RunProperties FontAt(int index)
    {
        // Font records are numbered from 0 but skip 4.
        int at = index > 4 ? index - 1 : index;
        return at >= 0 && at < _fonts.Count ? _fonts[at].Font : _fonts.Count > 0 ? _fonts[0].Font : new RunProperties { FontFamily = "Arial", FontSize = 10 };
    }

    private void ReadFormat(ReadOnlySpan<byte> d)
    {
        if (d.Length < 5) return;
        _formats[BinaryPrimitives.ReadUInt16LittleEndian(d)] = ReadString16(d, 2).Text;
    }

    private void ReadXf(ReadOnlySpan<byte> d)
    {
        if (d.Length < 20) { _xfs.Add(new Xf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)); return; }
        _xfs.Add(new Xf(BinaryPrimitives.ReadUInt16LittleEndian(d), BinaryPrimitives.ReadUInt16LittleEndian(d[2..]), d[6], d[7], d[8] & 0x0F,
            BinaryPrimitives.ReadUInt16LittleEndian(d[10..]), BinaryPrimitives.ReadUInt16LittleEndian(d[12..]), (int)BinaryPrimitives.ReadUInt32LittleEndian(d[14..]),
            (int)(BinaryPrimitives.ReadUInt32LittleEndian(d[14..]) >> 26), BinaryPrimitives.ReadUInt16LittleEndian(d[18..])));
    }

    private void ReadPalette(ReadOnlySpan<byte> d)
    {
        if (d.Length < 2) return;
        int count = Math.Min(BinaryPrimitives.ReadUInt16LittleEndian(d), (d.Length - 2) / 4);
        for (int i = 0; i < count && i < 56; i++) _palette[i] = new Color(d[2 + 4 * i], d[3 + 4 * i], d[4 + 4 * i]);
    }

    private Color? ColorOf(int index, Color? auto)
    {
        index &= 0x7FFF;
        if (index is >= 0 and <= 7) return _palette[index];
        if (index is >= 8 and <= 63) return _palette[index - 8];
        return index == 65 ? Color.White : auto; // 64 and 0x7FFF are the system's automatic colors
    }

    private static Border BorderOf(int style, Color? color)
    {
        if (style == 0) return Border.None;
        var (kind, width) = style switch
        {
            1 => (BorderStyle.Single, 0.75), 2 => (BorderStyle.Single, 1.5), 3 => (BorderStyle.Dashed, 0.75), 4 => (BorderStyle.Dotted, 0.75),
            5 => (BorderStyle.Thick, 2.25), 6 => (BorderStyle.Double, 2.25), 7 => (BorderStyle.Single, 0.25),
            8 or 10 or 12 => (BorderStyle.Dashed, 1.5), _ => (BorderStyle.Dashed, 0.75),
        };
        return new Border(kind, width, color ?? Color.Black);
    }

    private CellStyle StyleOf(int xfIndex)
    {
        if (_styleCache.TryGetValue(xfIndex, out var cached)) return cached;
        if (xfIndex < 0 || xfIndex >= _xfs.Count) return CellStyle.Default with { Font = FontAt(0) };
        var xf = _xfs[xfIndex];
        var font = FontAt(xf.Font);
        font = font with { Color = font.Color ?? Color.Black };
        int icvLeft = xf.BorderColors & 0x7F, icvRight = (xf.BorderColors >> 7) & 0x7F, icvTop = xf.TopBottomColors & 0x7F, icvBottom = (xf.TopBottomColors >> 7) & 0x7F;
        int icvFore = xf.FillColors & 0x7F, icvBack = (xf.FillColors >> 7) & 0x7F;
        Color? fill = null;
        if (xf.Fill == 1) fill = ColorOf(icvFore, Color.Black);
        else if (xf.Fill > 1 && ColorOf(icvFore, Color.Black) is { } fore)
            fill = ColorOf(icvBack, Color.White) is { } back ? back.Blend(fore, xf.Fill is 2 ? 0.5 : xf.Fill is 3 ? 0.75 : 0.25) : fore;
        int rotation = xf.Rotation;
        var style = new CellStyle
        {
            Font = font,
            Fill = fill,
            Borders = new BorderSet
            {
                Left = BorderOf(xf.Borders & 0xF, ColorOf(icvLeft, null)), Right = BorderOf((xf.Borders >> 4) & 0xF, ColorOf(icvRight, null)),
                Top = BorderOf((xf.Borders >> 8) & 0xF, ColorOf(icvTop, null)), Bottom = BorderOf((xf.Borders >> 12) & 0xF, ColorOf(icvBottom, null)),
            },
            HorizontalAlignment = (xf.Align & 7) switch { 1 => HorizontalAlignment.Left, 2 or 6 => HorizontalAlignment.Center, 3 => HorizontalAlignment.Right, 4 => HorizontalAlignment.Fill, 5 => HorizontalAlignment.Justify, _ => HorizontalAlignment.General },
            VerticalAlignment = ((xf.Align >> 4) & 7) switch { 0 => CellVerticalAlignment.Top, 1 => CellVerticalAlignment.Center, _ => CellVerticalAlignment.Bottom },
            Wrap = (xf.Align & 8) != 0,
            Indent = xf.Indent,
            Rotation = rotation is >= 0 and <= 90 ? rotation : rotation is > 90 and <= 180 ? 90 - rotation : 0,
            Stacked = rotation == 255,
        };
        return _styleCache[xfIndex] = style;
    }

    private string FormatCode(int xfIndex)
    {
        if (xfIndex < 0 || xfIndex >= _xfs.Count) return "General";
        int id = _xfs[xfIndex].Format;
        if (_formats.TryGetValue(id, out var code)) return code;
        return NumberFormat.BuiltIn(id) ?? "General";
    }

    // ------------------------------------------------------------------ sheets

    private void PutNumber(Worksheet sheet, int row, int column, int xf, double value)
    {
        var style = StyleOf(xf);
        var shown = NumberFormat.Parse(FormatCode(xf)).FormatNumber(value, _date1904);
        Put(sheet, row, column, new Cell
        {
            Text = shown.Text, Kind = CellValueKind.Number, Value = value,
            Style = shown.Color is { } c ? style with { Font = style.Font with { Color = c } } : style,
        });
    }

    private void Put(Worksheet sheet, int row, int column, Cell cell)
    {
        if (row >= MaxRows || column >= MaxColumns) return;
        if (++_cells > options.MaxCellCount) throw new InvalidDataException("The workbook has too many cells.");
        sheet.Cells[(row, column)] = cell;
    }

    private void PutText(Worksheet sheet, int row, int column, int xf, string text)
    {
        if (text.Length > MaxCellText) text = text[..MaxCellText];
        Put(sheet, row, column, new Cell { Text = text, Kind = CellValueKind.Text, Style = StyleOf(xf), IsEmptyString = text.Length == 0 });
    }

    private static double Rk(uint rk)
    {
        double v = (rk & 2) != 0 ? (int)rk >> 2 : BitConverter.Int64BitsToDouble((long)(rk & 0xFFFFFFFC) << 32);
        return (rk & 1) != 0 ? v / 100 : v;
    }

    private static string ErrorText(int code) => code switch { 0x00 => "#NULL!", 0x07 => "#DIV/0!", 0x0F => "#VALUE!", 0x17 => "#REF!", 0x1D => "#NAME?", 0x24 => "#NUM!", 0x2A => "#N/A", _ => "#VALUE!" };

    private void ReadSheet(Worksheet sheet, int bof)
    {
        int depth = 0;
        (int Row, int Column, int Xf)? pendingFormula = null;
        var setup = sheet.PageSetup;
        bool landscape = false, fitToPage = false;
        int paper = 1, fitW = 1, fitH = 1, scale = 100;
        string? header = null, footer = null;
        double? left = null, right = null, top = null, bottom = null, headerMargin = null, footerMargin = null;
        var drawing = new Drawing();
        for (int i = bof; i < _records.Count; i++)
        {
            var r = _records[i];
            var d = Data(r);
            if (r.Id is 0x0809 or 0x0009) { if (++depth > 1) { i = SkipSubstream(i); depth--; } continue; }
            if (r.Id == 0x000A) { if (--depth <= 0) break; continue; }
            switch (r.Id)
            {
                case 0x0203 when d.Length >= 14: PutNumber(sheet, U16(d, 0), U16(d, 2), U16(d, 4), BitConverter.ToDouble(d[6..14])); break;
                case 0x027E when d.Length >= 10: PutNumber(sheet, U16(d, 0), U16(d, 2), U16(d, 4), Rk(BinaryPrimitives.ReadUInt32LittleEndian(d[6..]))); break;
                case 0x00BD when d.Length >= 6:
                {
                    int row = U16(d, 0), first = U16(d, 2), lastColumn = U16(d, d.Length - 2);
                    for (int c = first, p = 4; c <= lastColumn && p + 6 <= d.Length - 2; c++, p += 6)
                        PutNumber(sheet, row, c, U16(d, p), Rk(BinaryPrimitives.ReadUInt32LittleEndian(d[(p + 2)..])));
                    break;
                }
                case 0x00FD when d.Length >= 10:
                {
                    uint index = BinaryPrimitives.ReadUInt32LittleEndian(d[6..]);
                    PutText(sheet, U16(d, 0), U16(d, 2), U16(d, 4), index < _sst.Count ? _sst[(int)index] : "");
                    break;
                }
                case 0x0204 when d.Length >= 9: PutText(sheet, U16(d, 0), U16(d, 2), U16(d, 4), ReadString16(d, 6).Text); break;
                case 0x0205 when d.Length >= 8:
                {
                    int row = U16(d, 0), column = U16(d, 2), xf = U16(d, 4);
                    if (d[7] == 0) Put(sheet, row, column, new Cell { Text = d[6] != 0 ? "TRUE" : "FALSE", Kind = CellValueKind.Bool, Style = StyleOf(xf) });
                    else Put(sheet, row, column, new Cell { Text = ErrorText(d[6]), Kind = CellValueKind.Error, Style = StyleOf(xf) });
                    break;
                }
                case 0x0201 when d.Length >= 6: if (StyleOf(U16(d, 4)) is { } s && (s.Fill is not null || s.Borders.HasAny)) Put(sheet, U16(d, 0), U16(d, 2), new Cell { Style = s }); break;
                case 0x00BE when d.Length >= 6:
                {
                    int row = U16(d, 0), first = U16(d, 2), lastColumn = U16(d, d.Length - 2);
                    for (int c = first, p = 4; c <= lastColumn && p + 2 <= d.Length - 2; c++, p += 2)
                        if (StyleOf(U16(d, p)) is { } s2 && (s2.Fill is not null || s2.Borders.HasAny)) Put(sheet, row, c, new Cell { Style = s2 });
                    break;
                }
                case 0x0006 when d.Length >= 20:
                {
                    // Row, column, XF, then the cached result in 8 bytes: a double, or 0xFFFF in the last two marks a string (type 0), boolean (1), error (2) or empty string (3).
                    int row = U16(d, 0), column = U16(d, 2), xf = U16(d, 4);
                    pendingFormula = null;
                    if (d[12] == 0xFF && d[13] == 0xFF)
                    {
                        switch (d[6])
                        {
                            case 0: pendingFormula = (row, column, xf); break; // a STRING record follows
                            case 1: Put(sheet, row, column, new Cell { Text = d[8] != 0 ? "TRUE" : "FALSE", Kind = CellValueKind.Bool, Style = StyleOf(xf) }); break;
                            case 2: Put(sheet, row, column, new Cell { Text = ErrorText(d[8]), Kind = CellValueKind.Error, Style = StyleOf(xf) }); break;
                            case 3: PutText(sheet, row, column, xf, ""); break;
                        }
                    }
                    else PutNumber(sheet, row, column, xf, BitConverter.ToDouble(d[6..14]));
                    break;
                }
                case 0x0207 when pendingFormula is { } p && d.Length >= 3:
                    PutText(sheet, p.Row, p.Column, p.Xf, ReadString16(d, 0).Text);
                    pendingFormula = null;
                    break;
                case 0x0208 when d.Length >= 16:
                {
                    int row = U16(d, 0), height = U16(d, 6), flags = (int)BinaryPrimitives.ReadUInt32LittleEndian(d[12..]);
                    if ((height & 0x8000) == 0 && height > 0 && row < MaxRows) sheet.RowHeights[row] = (height & 0x7FFF) / 20.0;
                    if ((flags & 0x20) != 0 && row < MaxRows) sheet.HiddenRows.Add(row);
                    break;
                }
                case 0x007D when d.Length >= 12:
                {
                    int first = U16(d, 0), last = Math.Min(U16(d, 2), MaxColumns - 1), width = U16(d, 4), flags = U16(d, 8);
                    for (int c = first; c <= last; c++)
                    {
                        if ((flags & 1) != 0) sheet.HiddenColumns.Add(c);
                        else sheet.ColumnWidths[c] = width / 256.0 * 7 * 0.75;
                    }
                    break;
                }
                case 0x0055 when d.Length >= 2: sheet.DefaultColumnWidth = (U16(d, 0) * 7 + 5) * 0.75; break;
                case 0x0225 when d.Length >= 4: if (U16(d, 2) > 0) sheet.DefaultRowHeight = U16(d, 2) / 20.0; break;
                case 0x00E5 when d.Length >= 2:
                {
                    int count = Math.Min(U16(d, 0), (d.Length - 2) / 8);
                    for (int k = 0; k < count; k++)
                    {
                        int p = 2 + 8 * k;
                        int r1 = U16(d, p), r2 = U16(d, p + 2), c1 = U16(d, p + 4), c2 = U16(d, p + 6);
                        if (r2 >= r1 && c2 >= c1 && (r2 > r1 || c2 > c1)) sheet.Merges.Add(new CellRange(r1, c1, Math.Min(r2, MaxRows - 1), Math.Min(c2, MaxColumns - 1)));
                    }
                    break;
                }
                case 0x00A1 when d.Length >= 34:
                    // fNoPls (0x04): the printer fields (paper, scale, orientation) were never set and hold junk; Excel uses its defaults.
                    fitW = U16(d, 6); fitH = U16(d, 8);
                    if ((U16(d, 10) & 4) == 0)
                    {
                        paper = U16(d, 0); scale = U16(d, 2);
                        landscape = (U16(d, 10) & 2) == 0 && (U16(d, 10) & 0x40) == 0;
                    }
                    headerMargin = BitConverter.ToDouble(d[16..24]); footerMargin = BitConverter.ToDouble(d[24..32]);
                    if ((U16(d, 10) & 1) != 0) setup.PageOrderOverThenDown = true;
                    break;
                case 0x00EC: i = AddEscher(drawing, i); break;
                case 0x005D when d.Length >= 10 && U16(d, 0) == 0x15: drawing.Objects.Add((U16(d, 4), U16(d, 8))); break;
                case 0x01B6 when d.Length >= 14: i = AddText(drawing, i); break;
                case 0x0081 when d.Length >= 2: fitToPage = (U16(d, 0) & 0x100) != 0; break;
                case 0x0014 when d.Length >= 3: header = d.Length > 2 ? ReadString16Short(d) : ""; break;
                case 0x0015 when d.Length >= 3: footer = d.Length > 2 ? ReadString16Short(d) : ""; break;
                case 0x0026 when d.Length >= 8: left = BitConverter.ToDouble(d[..8]); break;
                case 0x0027 when d.Length >= 8: right = BitConverter.ToDouble(d[..8]); break;
                case 0x0028 when d.Length >= 8: top = BitConverter.ToDouble(d[..8]); break;
                case 0x0029 when d.Length >= 8: bottom = BitConverter.ToDouble(d[..8]); break;
                case 0x002B when d.Length >= 2: setup.PrintGridLines = U16(d, 0) != 0; break;
                case 0x0083 when d.Length >= 2: setup.CenterHorizontally = U16(d, 0) != 0; break;
                case 0x0084 when d.Length >= 2: setup.CenterVertically = U16(d, 0) != 0; break;
                case 0x003D when d.Length >= 10: break; // WINDOW1 lives in the globals; nothing to read per sheet
                case 0x023E when d.Length >= 2: sheet.ShowGridLines = (U16(d, 0) & 0x0002) != 0; break; // WINDOW2 option flags: bit 1 shows gridlines
            }
        }

        AddDrawings(sheet, drawing);

        var (pageW, pageH) = PaperSize(paper);
        setup.Orientation = landscape ? Orientation.Landscape : Orientation.Portrait;
        (setup.PageWidth, setup.PageHeight) = landscape ? (pageH, pageW) : (pageW, pageH);
        if (fitToPage) { setup.FitToWidth = Math.Max(fitW, 0); setup.FitToHeight = Math.Max(fitH, 0); }
        else if (scale is >= 10 and <= 400) setup.Scale = scale / 100.0;
        double In(double? v, double fallback) => v is { } x && double.IsFinite(x) && x >= 0 && x < 20 ? x * 72 : fallback;
        setup.Margins = new Thickness(In(top, 72), In(right, 54), In(bottom, 72), In(left, 54));
        setup.HeaderDistance = In(headerMargin, 21.6);
        setup.FooterDistance = In(footerMargin, 21.6);
        if (header is not null) setup.Header = new HeaderFooterText { Odd = Split(header) };
        if (footer is not null) setup.Footer = new HeaderFooterText { Odd = Split(footer) };
    }

    private static int U16(ReadOnlySpan<byte> d, int at) => BinaryPrimitives.ReadUInt16LittleEndian(d[at..]);

    private static string ReadString16Short(ReadOnlySpan<byte> d)
    {
        int cch = U16(d, 0);
        if (cch == 0 || d.Length < 3) return "";
        return ReadChars(d, 3, cch, (d[2] & 1) != 0).Text;
    }

    private static HeaderFooterParts Split(string text)
    {
        var parts = new[] { new StringBuilder(), new StringBuilder(), new StringBuilder() };
        int current = 1;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c != '&' || i + 1 >= text.Length) { parts[current].Append(c); continue; }
            char next = text[i + 1];
            switch (next)
            {
                case 'L': current = 0; i++; break;
                case 'C': current = 1; i++; break;
                case 'R': current = 2; i++; break;
                default: parts[current].Append(c).Append(next); i++; break;
            }
        }
        return new HeaderFooterParts(parts[0].ToString(), parts[1].ToString(), parts[2].ToString());
    }

    private static (double, double) PaperSize(int id)
    {
        static (double, double) In(double w, double h) => (w * 72, h * 72);
        static (double, double) Mm(double w, double h) => (w * 72 / 25.4, h * 72 / 25.4);
        return id switch
        {
            3 => In(11, 17), 4 => In(17, 11), 5 => In(8.5, 14), 6 => In(5.5, 8.5), 7 => In(7.25, 10.5), 8 => Mm(297, 420), 9 or 10 => Mm(210, 297),
            11 => Mm(148, 210), 12 => Mm(257, 364), 13 => Mm(182, 257), 14 => In(8.5, 13), 15 => Mm(215, 275), 16 => In(10, 14),
            _ => In(8.5, 11),
        };
    }

    private static readonly Color[] DefaultPalette = new uint[]
    {
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
        0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
        0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
        0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
    }.Select(v => new Color((byte)(v >> 16), (byte)(v >> 8), (byte)v)).ToArray();
}
