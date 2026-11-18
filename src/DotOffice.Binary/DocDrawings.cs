using System.Buffers.Binary;
using DotOffice.Core;

namespace DotOffice.Binary;

/// <summary>
/// The floating objects of a .doc: shapes anchored at a 0x08 character, placed by an FSPA (PlcSpaMom for the main story, PlcSpaHdr for
/// the headers) and described by the Escher shape with the same id in the table stream's drawing (OfficeArtContent). Pictures come from the
/// drawing group's BLIP store; text boxes take their text from the text box story (PlcftxbxTxt).
/// </summary>
internal sealed partial class DocContext
{
    private const int MaxDrawingDepth = 4;

    private readonly record struct Fspa(int Spid, int Left, int Top, int Right, int Bottom, int Flags);

    private readonly Dictionary<int, Fspa> _mainSpas = [], _headerSpas = [];
    private readonly Dictionary<int, OfficeArt.EscherShape> _shapes = [];
    private readonly List<(int Start, int End)> _textBoxes = [], _headerTextBoxes = [];
    private List<ImageData?> _pictures = [];
    private int _drawingDepth;

    private void ReadDrawings()
    {
        int fc = FcLcb(50, out int lcb);
        if (lcb > 8 && fc >= 0 && (long)fc + lcb <= table.Length)
        {
            var escher = table.AsSpan(fc, lcb).ToArray();
            _pictures = OfficeArt.ReadStore(escher, main);
            // OfficeArtContent: the drawing group, then each drawing (main document, headers) as a dgglbl byte and its DgContainer.
            int at = 0;
            foreach (var (type, _, body, length) in OfficeArt.Records(escher, 0, escher.Length)) { if (type == 0xF000) at = body + length; break; }
            for (int k = 0; k < 2 && at + 9 <= escher.Length; k++)
            {
                foreach (var (type, _, body, length) in OfficeArt.Records(escher, at + 1, escher.Length))
                {
                    if (type == 0xF002) Collect(escher, body, body + length, 0);
                    at = body + length;
                    break;
                }
            }
        }
        ReadSpas(40, _mainSpas);
        ReadSpas(41, _headerSpas);

        // The text box stories follow the endnotes; the header text box stories follow them.
        int baseCp = _textLength + Math.Max(Lw(4), 0) + Math.Max(Lw(5), 0) + Math.Max(Lw(6), 0) + Math.Max(Lw(7), 0) + Math.Max(Lw(8), 0);
        ReadTextBoxes(56, baseCp, _textBoxes);
        ReadTextBoxes(58, baseCp + Math.Max(Lw(9), 0), _headerTextBoxes);
    }

    private void ReadTextBoxes(int pair, int baseCp, List<(int, int)> into)
    {
        int fc = FcLcb(pair, out int lcb);
        if (lcb < 4 + 26 || fc < 0 || (long)fc + lcb > table.Length) return;
        var plc = table.AsSpan(fc, lcb);
        int n = (lcb - 4) / 26;
        for (int i = 0; i < n; i++) into.Add((baseCp + (int)U32(plc, 4 * i), baseCp + (int)U32(plc, 4 * (i + 1))));
    }

    private void Collect(byte[] escher, int start, int end, int depth)
    {
        if (depth > 32) return;
        foreach (var (type, _, body, length) in OfficeArt.Records(escher, start, end))
        {
            if (type == 0xF003) Collect(escher, body, body + length, depth + 1);
            else if (type == 0xF004 && _shapes.Count < 100_000)
            {
                var shape = OfficeArt.ReadShape(escher, body, body + length);
                _shapes.TryAdd(shape.Spid, shape);
            }
        }
    }

    private void ReadSpas(int pair, Dictionary<int, Fspa> into)
    {
        int fc = FcLcb(pair, out int lcb);
        if (lcb < 4 + 26 || fc < 0 || (long)fc + lcb > table.Length) return;
        var plc = table.AsSpan(fc, lcb);
        int n = (lcb - 4) / 30;
        for (int i = 0; i < n; i++)
        {
            int at = 4 * (n + 1) + 26 * i;
            var f = plc[at..];
            into.TryAdd((int)U32(plc, 4 * i), new Fspa(BinaryPrimitives.ReadInt32LittleEndian(f), BinaryPrimitives.ReadInt32LittleEndian(f[4..]),
                BinaryPrimitives.ReadInt32LittleEndian(f[8..]), BinaryPrimitives.ReadInt32LittleEndian(f[12..]), BinaryPrimitives.ReadInt32LittleEndian(f[16..]), U16(f, 20)));
        }
    }

    /// <summary>The floating picture or text box anchored at <paramref name="cp"/>, or null when there is none (or it is not drawn).</summary>
    private Inline? FloatingAt(int cp, RunProperties props)
    {
        bool header = cp >= _textLength;
        int headerBase = _textLength + Math.Max(Lw(4), 0);
        if (!(header ? _headerSpas.TryGetValue(cp - headerBase, out var spa) : _mainSpas.TryGetValue(cp, out spa))) return null;
        if (!_shapes.TryGetValue(spa.Spid, out var shape) || shape.Group) return null;
        var p = shape.Properties;
        uint Prop(int id, uint fallback) => p.TryGetValue(id, out var v) ? v : fallback;
        bool Bool(int id, uint bit, bool fallback) => p.TryGetValue(id, out var v) && (v & (bit << 16)) != 0 ? (v & bit) != 0 : fallback;
        double Emu(int id, double fallback) => p.TryGetValue(id, out var v) && v < 20_000_000 ? v / 12700.0 : fallback;

        double width = (spa.Right - spa.Left) / 20.0, height = (spa.Bottom - spa.Top) / 20.0;
        if (!(width > 0 && height > 0 && width < 20_000 && height < 20_000)) return null;
        // Word 2000 and later put the position in the shape (posh, posrelh, posv, posrelv); the FSPA's bx and by are the older form.
        var horizontalTo = p.TryGetValue(0x390, out var rh)
            ? rh switch { 0 => FloatReference.Margin, 1 => FloatReference.Page, 3 => FloatReference.Character, _ => FloatReference.Column }
            : ((spa.Flags >> 1) & 3) switch { 0 => FloatReference.Margin, 1 => FloatReference.Page, _ => FloatReference.Column };
        var verticalTo = p.TryGetValue(0x392, out var rv)
            ? rv switch { 0 => FloatReference.Margin, 1 => FloatReference.Page, 3 => FloatReference.Line, _ => FloatReference.Paragraph }
            : ((spa.Flags >> 3) & 3) switch { 0 => FloatReference.Margin, 1 => FloatReference.Page, _ => FloatReference.Paragraph };
        static FloatAlignment Align(uint v) => v switch { 1 => FloatAlignment.Start, 2 => FloatAlignment.Center, 3 => FloatAlignment.End, 4 => FloatAlignment.Inside, 5 => FloatAlignment.Outside, _ => FloatAlignment.Offset };
        var horizontal = new FloatPosition(horizontalTo, Align(Prop(0x38F, 0)), spa.Left / 20.0);
        var vertical = new FloatPosition(verticalTo, Align(Prop(0x391, 0)), spa.Top / 20.0);
        int wr = (spa.Flags >> 5) & 0xF;
        var wrap = wr switch { 1 => TextWrap.TopAndBottom, 3 => TextWrap.None, _ => TextWrap.Square };
        bool behind = (spa.Flags & 0x4000) != 0;

        if (p.TryGetValue(0x104, out var pib) && pib >= 1 && pib <= _pictures.Count && _pictures[(int)pib - 1] is { } picture)
            return new FloatingImage(picture, width, height, props) { Horizontal = horizontal, Vertical = vertical, Wrap = wrap, BehindText = behind };

        var geometry = shape.Type switch { 2 => ShapeGeometry.RoundRectangle, 3 => ShapeGeometry.Ellipse, 1 or 202 => ShapeGeometry.Rectangle, _ => (ShapeGeometry?)null };
        if (geometry is null) return null; // lines, arrows, freeforms and the rest are not drawn yet
        var box = new FloatingTextBox(width, height, props)
        {
            Horizontal = horizontal, Vertical = vertical, Wrap = wrap, BehindText = behind, Geometry = geometry.Value,
            Fill = Bool(0x1BF, 0x10, true) ? Rgb(Prop(0x181, 0xFFFFFF), new Color(255, 255, 255)) : null,
            Outline = Bool(0x1FF, 0x8, true) ? Rgb(Prop(0x1C0, 0), new Color(0, 0, 0)) : null,
            OutlineWidth = Emu(0x1CB, 0.75) is > 0 and var w ? w : 0.5, // a zero width is a hairline
            Insets = new Thickness(Emu(0x82, 3.6), Emu(0x83, 7.2), Emu(0x84, 3.6), Emu(0x81, 7.2)),
            VerticalAnchor = Prop(0x87, 0) switch { 1 or 4 or 7 => TextBoxAnchor.Center, 2 or 5 or 8 => TextBoxAnchor.Bottom, _ => TextBoxAnchor.Top },
            Rotation = p.TryGetValue(0x4, out var r) && double.IsFinite((int)r / 65536.0) ? (int)r / 65536.0 % 360 : 0,
        };
        // lTxid: the text box story's index (plus one) in its high word.
        if (p.TryGetValue(0x80, out var txid) && (int)(txid >> 16) - 1 is int story && story >= 0 && story < (header ? _headerTextBoxes : _textBoxes).Count && _drawingDepth < MaxDrawingDepth)
        {
            var (start, end) = (header ? _headerTextBoxes : _textBoxes)[story];
            if (start < end && end <= _text.Length)
            {
                _drawingDepth++;
                try { box.Blocks.AddRange(Build(start, end, story: true).Sections[0].Blocks); }
                finally { _drawingDepth--; }
            }
        }
        return box;
    }

    /// <summary>An Office Art color: RGB in the low three bytes; scheme, system and other indirect colors give the fallback.</summary>
    private static Color Rgb(uint value, Color fallback) =>
        (value & 0xFF000000) != 0 ? fallback : new Color((byte)value, (byte)(value >> 8), (byte)(value >> 16));
}
