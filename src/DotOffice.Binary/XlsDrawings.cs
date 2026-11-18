using System.Buffers.Binary;
using System.Text;
using DotOffice.Core;

namespace DotOffice.Binary;

/// <summary>
/// The drawing layer of an .xls sheet: the Escher (Office Art) records of its MSODRAWING records, the OBJ record of each shape and the
/// TXO record (with its CONTINUE records) of each shape with text. Text boxes, rectangles, ovals and lines become sheet shapes.
/// </summary>
internal sealed partial class XlsContext
{
    private const int MaxDrawingBytes = 64 << 20, MaxShapes = 10_000;

    private readonly List<ImageData?> _pictures = [];

    private sealed record TextBox(string Text, List<(int At, int Font)> Runs, int Options);

    private sealed class Drawing
    {
        public readonly List<byte> Escher = [];
        /// <summary>The OBJ records in order: object type and option flags.</summary>
        public readonly List<(int Type, int Flags)> Objects = [];
        public readonly List<TextBox> Texts = [];
    }

    private void AddDrawings(Worksheet sheet, Drawing drawing)
    {
        if (drawing.Escher.Count == 0) return;
        var escher = drawing.Escher.ToArray();
        int obj = 0, txo = 0, z = 0;
        Visit(0, escher.Length, null, 0);

        // A group's children are placed in the group's own coordinates (its Spgr rectangle) and take its anchor.
        void Visit(int start, int end, (Anchor Anchor, FrameFraction Frame, Rect Space)? group, int depth)
        {
            if (depth > 32) return;
            (Anchor, FrameFraction, Rect)? own = null;
            foreach (var (type, inst, body, length) in OfficeArt.Records(escher, start, end))
            {
                if (type == 0xF002) { Visit(body, body + length, group, depth + 1); continue; }
                if (type == 0xF003) { Visit(body, body + length, own ?? group, depth + 1); continue; }
                if (type != 0xF004) continue;
                var shape = ReadShape(escher, body, body + length);
                int? objType = null, objFlags = null;
                if (shape.HasClientData && obj < drawing.Objects.Count) (objType, objFlags) = drawing.Objects[obj++];
                TextBox? text = shape.HasTextbox && txo < drawing.Texts.Count ? drawing.Texts[txo++] : null;
                if (shape.Patriarch) continue;
                (Anchor Anchor, FrameFraction Frame)? place = shape.Anchor is { } a ? (a, FrameFraction.Whole)
                    : shape.ChildRect is { } c && group is { } g ? (g.Anchor, g.Frame.Within(Relative(c, g.Space))) : null;
                if (place is not { } p) continue;
                if (shape.Group) { if (own is null && shape.GroupRect is { } space) own = (p.Anchor, p.Frame, space); continue; }
                if (sheet.Shapes.Count >= MaxShapes || objFlags is { } f && (f & 0x10) == 0) continue; // fPrint off: not printed
                if ((objType == 8 || objType is null && shape.Type == 75) && PictureOf(shape) is { } picture)
                {
                    if (sheet.Images.Count < MaxShapes) sheet.Images.Add(Placed(new SheetImage { Image = picture, ZOrder = z++ }, sheet, shape, p.Anchor, p.Frame));
                    continue;
                }
                if (objType is not (1 or 2 or 3 or 6 or 7 or 0x0B or 0x0C or 0x0E) && !(objType is null && text is not null)) continue;
                sheet.Shapes.Add(Placed(BuildShape(sheet, shape, objType, text, z++), sheet, shape, p.Anchor, p.Frame));
            }
        }
    }

    private readonly record struct Anchor(int Col1, int Dx1, int Row1, int Dy1, int Col2, int Dx2, int Row2, int Dy2);
    private readonly record struct Rect(int Left, int Top, int Right, int Bottom);

    private sealed class EscherShape
    {
        public int Type;
        public bool Group, Patriarch, FlipH, FlipV, HasClientData, HasTextbox;
        public Anchor? Anchor;
        public Rect? ChildRect, GroupRect;
        public readonly Dictionary<int, uint> Properties = [];
    }

    private static FrameFraction Relative(Rect child, Rect space)
    {
        double w = space.Right - space.Left, h = space.Bottom - space.Top;
        if (w <= 0 || h <= 0) return FrameFraction.Whole;
        return new FrameFraction((child.Left - space.Left) / w, (child.Top - space.Top) / h, (child.Right - child.Left) / w, (child.Bottom - child.Top) / h);
    }

    private static EscherShape ReadShape(byte[] escher, int start, int end)
    {
        var shape = new EscherShape();
        Rect ReadRect(int at) => new(BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at)), BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at + 4)),
            BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at + 8)), BinaryPrimitives.ReadInt32LittleEndian(escher.AsSpan(at + 12)));
        foreach (var (type, inst, body, length) in OfficeArt.Records(escher, start, end))
        {
            switch (type)
            {
                case 0xF009 when length >= 16: shape.GroupRect = ReadRect(body); break;
                case 0xF00A when length >= 8:
                    uint flags = BinaryPrimitives.ReadUInt32LittleEndian(escher.AsSpan(body + 4));
                    shape.Type = inst;
                    shape.Group = (flags & 1) != 0; shape.Patriarch = (flags & 4) != 0;
                    shape.FlipH = (flags & 0x40) != 0; shape.FlipV = (flags & 0x80) != 0;
                    break;
                case 0xF00B:
                    for (int i = 0; i < inst && body + 6 * i + 6 <= body + length; i++)
                    {
                        int id = BinaryPrimitives.ReadUInt16LittleEndian(escher.AsSpan(body + 6 * i));
                        if ((id & 0x8000) == 0) shape.Properties[id & 0x3FFF] = BinaryPrimitives.ReadUInt32LittleEndian(escher.AsSpan(body + 6 * i + 2));
                    }
                    break;
                case 0xF00F when length >= 16: shape.ChildRect = ReadRect(body); break;
                case 0xF010 when length >= 18:
                    int U(int k) => BinaryPrimitives.ReadUInt16LittleEndian(escher.AsSpan(body + 2 + 2 * k));
                    shape.Anchor = new Anchor(U(0), U(1), U(2), U(3), U(4), U(5), U(6), U(7));
                    break;
                case 0xF011: shape.HasClientData = true; break;
                case 0xF00D: shape.HasTextbox = true; break;
            }
        }
        return shape;
    }

    private SheetShape BuildShape(Worksheet sheet, EscherShape shape, int? objType, TextBox? text, int z)
    {
        var props = shape.Properties;
        uint Prop(int id, uint fallback) => props.TryGetValue(id, out var v) ? v : fallback;
        bool Bool(int id, uint bit, bool fallback) => props.TryGetValue(id, out var v) && (v & (bit << 16)) != 0 ? (v & bit) != 0 : fallback;
        var geometry = objType switch
        {
            1 => ShapeGeometry.Line,
            3 => ShapeGeometry.Ellipse,
            _ => shape.Type switch { 2 => ShapeGeometry.RoundRectangle, 3 => ShapeGeometry.Ellipse, 20 => ShapeGeometry.Line, _ => ShapeGeometry.Rectangle },
        };
        // Form controls (buttons, check boxes, option buttons, labels) print their caption; a check box or option button leaves room for its box.
        bool control = objType is 7 or 0x0B or 0x0C or 0x0E;
        bool filled = geometry != ShapeGeometry.Line && (control ? objType == 7 : Bool(0x1BF, 0x10, true)), lined = control ? objType == 7 : Bool(0x1FF, 0x8, true);
        double Emu(int id, double fallback) => props.TryGetValue(id, out var v) && v < 20_000_000 ? v / 12700.0 : fallback;

        var paragraphs = new List<ShapeParagraph>();
        var align = HorizontalAlignment.Left;
        var anchor = ShapeTextAnchor.Top;
        if (text is not null)
        {
            align = ((text.Options >> 1) & 7) switch { 2 => HorizontalAlignment.Center, 3 => HorizontalAlignment.Right, 4 => HorizontalAlignment.Justify, _ => HorizontalAlignment.Left };
            anchor = ((text.Options >> 4) & 7) switch { 2 => ShapeTextAnchor.Middle, 3 => ShapeTextAnchor.Bottom, _ => ShapeTextAnchor.Top };
            paragraphs = Paragraphs(text, align);
        }
        return new SheetShape
        {
            Geometry = geometry,
            Fill = filled ? EscherColor(Prop(0x181, 0xFFFFFF), Color.White) : null,
            LineColor = lined ? EscherColor(Prop(0x1C0, 0), new Color(0, 0, 0)) : null,
            LineWidth = Emu(0x1CB, 0.75),
            LineStyle = Prop(0x1CE, 0) switch { 0 => StrokeStyle.Solid, 2 or 3 or 9 => StrokeStyle.Dotted, _ => StrokeStyle.Dashed },
            Paragraphs = paragraphs,
            TextAnchor = anchor,
            TextInsets = control ? new Thickness(1, 1, 1, objType is 0x0B or 0x0C ? 16 : 1)
                : new Thickness(Emu(0x82, 3.6), Emu(0x83, 7.2), Emu(0x84, 3.6), Emu(0x81, 7.2)),
            WrapText = Prop(0x85, 0) != 2,
            ZOrder = z,
        };
    }

    /// <summary>
    /// Places an object at its cell anchor. Offsets are 1/1024 of the column's width and 1/256 of the row's height. A shape turned by
    /// 45° to 135° (or 225° to 315°) keeps the anchor of its turned bounds, so its own frame is that box with width and height swapped.
    /// </summary>
    private static T Placed<T>(T item, Worksheet sheet, EscherShape shape, Anchor a, FrameFraction frame) where T : SheetObject
    {
        AnchorPoint Point(int col, int dx, int row, int dy)
        {
            col = Math.Clamp(col, 0, MaxColumns - 1); row = Math.Clamp(row, 0, MaxRows - 1);
            return new AnchorPoint(col, row, Math.Min(dx, 1024) / 1024.0 * sheet.ColumnWidth(col), Math.Min(dy, 256) / 256.0 * sheet.RowHeight(row));
        }
        var from = Point(a.Col1, a.Dx1, a.Row1, a.Dy1);
        var to = Point(a.Col2, a.Dx2, a.Row2, a.Dy2);
        double rotation = shape.Properties.TryGetValue(0x0004, out var r) ? (int)r / 65536.0 : 0;
        if (!double.IsFinite(rotation) || Math.Abs(rotation) > 36_000) rotation = 0;
        double turn = ((rotation % 360) + 360) % 360;
        if (turn is >= 45 and < 135 or >= 225 and < 315)
        {
            double w = Distance(from.Column, to.Column, sheet.ColumnWidth) + to.ColumnOffset - from.ColumnOffset;
            double h = Distance(from.Row, to.Row, sheet.RowHeight) + to.RowOffset - from.RowOffset;
            w *= frame.Width; h *= frame.Height;
            if (w > 0 && h > 0) frame = frame.Within(new FrameFraction((1 - h / w) / 2, (1 - w / h) / 2, h / w, w / h));
        }
        return item with
        {
            Anchor = ImageAnchorKind.TwoCell, From = from, To = to, Frame = frame,
            FlipH = shape.FlipH, FlipV = shape.FlipV,
            Rotation = -rotation, // clockwise in the file, counter-clockwise in the model
        };
    }

    private static double Distance(int first, int last, Func<int, double> size)
    {
        double total = 0;
        for (int i = first; i < last && i - first < 100_000; i++) total += size(i);
        return total;
    }

    private ImageData? PictureOf(EscherShape shape) =>
        shape.Properties.TryGetValue(0x0104, out var pib) && pib >= 1 && pib <= _pictures.Count ? _pictures[(int)pib - 1] : null;

    /// <summary>The MSODRAWINGGROUP record (with its CONTINUE records) in the workbook globals: the pictures the sheets' shapes refer to.</summary>
    private int ReadDrawingGroup(int index)
    {
        var escher = new List<byte>();
        int last = index;
        escher.AddRange(Data(_records[index]));
        while (last + 1 < _records.Count && _records[last + 1].Id is 0x003C or 0x00EB && escher.Count < MaxDrawingBytes) escher.AddRange(Data(_records[++last]));
        _pictures.AddRange(OfficeArt.ReadStore(escher.ToArray()));
        return last;
    }

    private List<ShapeParagraph> Paragraphs(TextBox text, HorizontalAlignment align)
    {
        var result = new List<ShapeParagraph>();
        var runs = text.Runs.Where(r => r.At < text.Text.Length).OrderBy(r => r.At).ToList();
        if (runs.Count == 0 || runs[0].At > 0) runs.Insert(0, (0, 0));
        RunProperties FontAtChar(int at)
        {
            int font = runs[0].Font;
            foreach (var r in runs) { if (r.At > at) break; font = r.Font; }
            return FontAt(font);
        }
        int start = 0;
        var body = text.Text.Replace("\r\n", "\n").Replace('\r', '\n');
        // Runs index the original text; with \r\n collapsed the indices shift, so only map them when nothing was collapsed.
        bool exact = body.Length == text.Text.Length;
        foreach (var line in body.Split('\n'))
        {
            var pieces = new List<CellTextRun>();
            int i = 0;
            while (i < line.Length)
            {
                var font = exact ? FontAtChar(start + i) : FontAtChar(0);
                int j = i + 1;
                while (j < line.Length && (!exact || ReferenceEquals(FontAtChar(start + j), font))) j++;
                pieces.Add(new CellTextRun(line[i..j], font));
                i = j;
            }
            double size = pieces.Count > 0 ? pieces.Max(p => p.Font.FontSize) : (exact ? FontAtChar(Math.Min(start, Math.Max(text.Text.Length - 1, 0))) : FontAtChar(0)).FontSize;
            result.Add(new ShapeParagraph(align, pieces, size));
            start += line.Length + 1;
        }
        return result;
    }

    private Color? EscherColor(uint value, Color fallback)
    {
        if ((value & 0x08000000) != 0) return ColorOf((int)(value & 0xFF), fallback);
        if ((value & 0xF0000000) != 0) return fallback; // system colors and other indirections
        return new Color((byte)value, (byte)(value >> 8), (byte)(value >> 16));
    }

    /// <summary>Appends an MSODRAWING record and the CONTINUE records after it to the sheet's Escher stream; returns the last record read.</summary>
    private int AddEscher(Drawing drawing, int index)
    {
        int last = index;
        void Add(Record r) { if (drawing.Escher.Count + r.Length <= MaxDrawingBytes) drawing.Escher.AddRange(data.AsSpan(r.Offset, r.Length)); }
        Add(_records[index]);
        while (last + 1 < _records.Count && _records[last + 1].Id == 0x003C) Add(_records[++last]);
        return last;
    }

    /// <summary>A TXO record: options, then the text in CONTINUE records (each starts with its own 8- or 16-bit flag), then the formatting runs.</summary>
    private int AddText(Drawing drawing, int index)
    {
        var d = Data(_records[index]);
        int options = U16(d, 0), cch = U16(d, 10), cbRuns = U16(d, 12);
        int last = index;
        var text = new StringBuilder();
        while (text.Length < cch && last + 1 < _records.Count && _records[last + 1].Id == 0x003C)
        {
            var c = Data(_records[++last]);
            if (c.Length < 1) continue;
            text.Append(ReadChars(c, 1, cch - text.Length, (c[0] & 1) != 0).Text);
        }
        var runBytes = new List<byte>();
        while (runBytes.Count < cbRuns && cch > 0 && last + 1 < _records.Count && _records[last + 1].Id == 0x003C)
            runBytes.AddRange(Data(_records[++last]));
        var runs = new List<(int, int)>();
        for (int k = 0; k + 4 <= runBytes.Count && runs.Count < 4096; k += 8)
            runs.Add((runBytes[k] | runBytes[k + 1] << 8, runBytes[k + 2] | runBytes[k + 3] << 8));
        if (drawing.Texts.Count < MaxShapes) drawing.Texts.Add(new TextBox(text.ToString(), runs, options));
        return last;
    }
}
