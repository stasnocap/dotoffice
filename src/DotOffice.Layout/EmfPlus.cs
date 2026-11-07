using System.Buffers.Binary;
using System.Text;
using DotOffice.Core;
using DotOffice.Fonts;

namespace DotOffice.Layout;

/// <summary>
/// A GDI+ brush: a solid color, or for gradients also <see cref="Gradient"/> (fills draw it; pens and hatches use the average color).
/// </summary>
internal sealed record PlusBrush(Color Color, double Alpha)
{
    public PlusGradient? Gradient { get; init; }
}

/// <summary>
/// A gradient brush in brush space: linear ones run from the left to the right edge of <see cref="Rect"/>; path ones (<see cref="Radial"/>)
/// blend from the stop at 0 at <see cref="Center"/> to the last stop at the bounds of their boundary. <see cref="Transform"/> maps brush space
/// to world space.
/// </summary>
internal sealed record PlusGradient(bool Radial, (double X, double Y, double W, double H) Rect, (double X, double Y) Center, Affine Transform,
    IReadOnlyList<GradientStop> Stops)
{
    /// <summary>The GDI+ wrap mode: 0 tile, 1 tile flipping every other copy (2 and 3 flip on the other axis), 4 clamp to the end colors.</summary>
    public int Wrap { get; init; }

    /// <summary>For path gradients: the boundary is an upright rectangle, so the colors run in concentric rectangles.</summary>
    public bool Rectangular { get; init; }
}

internal sealed record PlusPen(double Width, int Unit, int LineStyle, PlusBrush Brush);

internal sealed record PlusFont(double EmSize, int Unit, int Style, string Family);

internal sealed record PlusStringFormat(int Flags, int Alignment, int LineAlignment);

internal sealed record PlusImage(ImageData? Data, int Width, int Height);

/// <summary>
/// A region as the intersection of outlines in world coordinates (each an even-odd set of rings separated by NaN points); empty is
/// everything (an infinite region).
/// </summary>
internal sealed record PlusRegion(List<List<(double X, double Y)>> Outlines, bool Empty = false);

/// <summary>
/// Plays EMF+ records ([MS-EMFPLUS]) into the same output as the <see cref="GdiPlayer"/> of their EMF, sharing its clip groups and its
/// device-to-box mapping. Between <c>GetDC</c> and the next EMF+ record the GDI records draw (<see cref="GdiActive"/>).
/// </summary>
internal sealed class EmfPlusPlayer(GdiPlayer gdi, FontSelector fonts)
{
    private sealed class State
    {
        public Affine World = Affine.Identity;
        public int PageUnit = 2;
        public double PageScale = 1;
        public IReadOnlyList<(double X, double Y)>[] Clip = [];

        public State Copy() => (State)MemberwiseClone();
    }

    private State _state = new();
    private readonly Dictionary<int, State> _saved = [];
    private readonly Dictionary<int, object> _objects = [];
    private readonly Dictionary<int, (int Total, MemoryStream Data)> _partial = [];
    private double _dpiX = 96, _dpiY = 96;

    /// <summary>Whether GDI records draw: after <c>GetDC</c>, until the next EMF+ record.</summary>
    public bool GdiActive { get; private set; }

    /// <summary>Whether any record drew.</summary>
    public bool Drew { get; private set; }

    // ---- Coordinates ----

    /// <summary>Device units per unit (UnitType: world 0, display 1, pixel 2, point 3, inch 4, document 5, millimetre 6).</summary>
    private double UnitScale(int unit, bool vertical)
    {
        double dpi = vertical ? _dpiY : _dpiX;
        return unit switch { 3 => dpi / 72, 4 => dpi, 5 => dpi / 300, 6 => dpi / 25.4, _ => 1 };
    }

    /// <summary>EMF+ device pixels (at the header's resolution) to the output box, through the EMF's reference device.</summary>
    private Affine DeviceToBox =>
        new Affine(25.4 / _dpiX * gdi.PixelsPerMm.X, 0, 0, 25.4 / _dpiY * gdi.PixelsPerMm.Y, 0, 0).Then(gdi.DeviceToBox);

    private Affine PageToDevice =>
        new(_state.PageScale * UnitScale(_state.PageUnit, false), 0, 0, _state.PageScale * UnitScale(_state.PageUnit, true), 0, 0);

    private Affine Transform => _state.World.Then(PageToDevice).Then(DeviceToBox);

    /// <summary>A length in <paramref name="unit"/> as world units.</summary>
    private double ToWorld(double length, int unit) =>
        unit == 0 ? length : length * UnitScale(unit, true) / (_state.PageScale * UnitScale(_state.PageUnit, true));

    private void Add(LayoutItem item)
    {
        gdi.Groups.Add(item, _state.Clip);
        Drew = true;
    }

    // ---- Records ----

    /// <summary>Plays the EMF+ records held by one EMF comment record (the bytes after its "EMF+" identifier).</summary>
    public void Play(ReadOnlySpan<byte> d)
    {
        int pos = 0;
        while (pos + 12 <= d.Length)
        {
            int type = U16(d, pos), flags = U16(d, pos + 2);
            int size = I32(d, pos + 4), dataSize = I32(d, pos + 8);
            if (size < 12 || pos + size > d.Length || dataSize < 0 || dataSize > size - 12) return;
            var r = d.Slice(pos + 12, dataSize);
            pos += size;
            if (type == 0x4002) return;
            GdiActive = false;
            Record(type, flags, r);
        }
    }

    private void Record(int type, int flags, ReadOnlySpan<byte> r)
    {
        int id = flags & 0xFF;
        bool compressed = (flags & 0x4000) != 0;
        switch (type)
        {
            case 0x4001 when r.Length >= 16: // Header
                if (I32(r, 8) > 0) _dpiX = I32(r, 8);
                if (I32(r, 12) > 0) _dpiY = I32(r, 12);
                break;
            case 0x4004: // GetDC
                GdiActive = true;
                break;
            case 0x4008: // Object
                ObjectRecord(flags, r);
                break;
            case 0x4009 when r.Length >= 4: // Clear
            {
                var b = ReadColor(r, 0);
                double h = gdi.BoxHeight;
                Fill(RectPath(0, 0, Metafiles.UnitBox, h), b, box: true);
                break;
            }
            case 0x400A when r.Length >= 8: // FillRects
            {
                var brush = BrushOf(flags, I32(r, 0));
                int count = I32(r, 4);
                foreach (var (x, y, w, h) in Rects(r, 8, count, compressed)) Fill(RectPath(x, y, w, h), brush);
                break;
            }
            case 0x400B when r.Length >= 4: // DrawRects
            {
                if (PenOf(id) is not { } pen) break;
                foreach (var (x, y, w, h) in Rects(r, 4, I32(r, 0), compressed)) Stroke(RectPath(x, y, w, h), pen);
                break;
            }
            case 0x400C when r.Length >= 8: // FillPolygon
                Fill(Polyline(Points(r, 8, I32(r, 4), compressed, (flags & 0x0800) != 0), true), BrushOf(flags, I32(r, 0)));
                break;
            case 0x400D when r.Length >= 4: // DrawLines
                if (PenOf(id) is { } linePen) Stroke(Polyline(Points(r, 4, I32(r, 0), compressed, (flags & 0x0800) != 0), (flags & 0x2000) != 0), linePen);
                break;
            case 0x400E when r.Length >= 4: // FillEllipse
            {
                var (x, y, w, h) = Rects(r, 4, 1, compressed)[0];
                Fill(Ellipse(x, y, w, h), BrushOf(flags, I32(r, 0)));
                break;
            }
            case 0x400F: // DrawEllipse
            {
                if (PenOf(id) is not { } pen) break;
                var (x, y, w, h) = Rects(r, 0, 1, compressed)[0];
                Stroke(Ellipse(x, y, w, h), pen);
                break;
            }
            case 0x4010 when r.Length >= 12: // FillPie
            {
                var (x, y, w, h) = Rects(r, 12, 1, compressed)[0];
                Fill(Arc(x, y, w, h, F32(r, 4), F32(r, 8), pie: true), BrushOf(flags, I32(r, 0)));
                break;
            }
            case 0x4011 or 0x4012 when r.Length >= 8: // DrawPie, DrawArc
            {
                if (PenOf(id) is not { } pen) break;
                var (x, y, w, h) = Rects(r, 8, 1, compressed)[0];
                Stroke(Arc(x, y, w, h, F32(r, 0), F32(r, 4), pie: type == 0x4011), pen);
                break;
            }
            case 0x4013 when r.Length >= 4: // FillRegion
                if (_objects.GetValueOrDefault(id) is PlusRegion region) FillRegion(region, BrushOf(flags, I32(r, 0)));
                break;
            case 0x4014 when r.Length >= 4: // FillPath
                if (_objects.GetValueOrDefault(id) is List<PathPoint> fillPath) Fill(fillPath, BrushOf(flags, I32(r, 0)));
                break;
            case 0x4015 when r.Length >= 4: // DrawPath
                if (_objects.GetValueOrDefault(id) is List<PathPoint> drawPath && PenOf(I32(r, 0)) is { } pathPen) Stroke(drawPath, pathPen);
                break;
            case 0x4016 when r.Length >= 12: // FillClosedCurve
                Fill(Curve(Points(r, 12, I32(r, 8), compressed, (flags & 0x0800) != 0), F32(r, 4), true), BrushOf(flags, I32(r, 0)), (flags & 0x2000) == 0);
                break;
            case 0x4017 when r.Length >= 8: // DrawClosedCurve
                if (PenOf(id) is { } closedPen) Stroke(Curve(Points(r, 8, I32(r, 4), compressed, (flags & 0x0800) != 0), F32(r, 0), true), closedPen);
                break;
            case 0x4018 when r.Length >= 16: // DrawCurve
            {
                if (PenOf(id) is not { } pen) break;
                var points = Points(r, 16, I32(r, 12), compressed, false);
                int offset = I32(r, 4), segments = I32(r, 8);
                if (offset < 0 || segments < 1 || offset + segments >= points.Count) break;
                // The whole curve's tangents, drawn over the chosen segments.
                var all = Curve(points, F32(r, 0), false);
                var part = new List<PathPoint> { all[3 * offset] with { Kind = PathPointKind.Move } };
                part.AddRange(all.GetRange(3 * offset + 1, 3 * segments));
                Stroke(part, pen);
                break;
            }
            case 0x4019 when r.Length >= 4: // DrawBeziers
            {
                if (PenOf(id) is not { } pen) break;
                var points = Points(r, 4, I32(r, 0), compressed, (flags & 0x0800) != 0);
                if (points.Count < 4) break;
                var path = new List<PathPoint> { P(points[0].X, points[0].Y, PathPointKind.Move) };
                for (int k = 1; k + 2 < points.Count; k += 3)
                    for (int j = 0; j < 3; j++) path.Add(P(points[k + j].X, points[k + j].Y, PathPointKind.Curve));
                Stroke(path, pen);
                break;
            }
            case 0x401A when r.Length >= 24: // DrawImage
            {
                if (_objects.GetValueOrDefault(id) is not PlusImage image) break;
                var src = (F32(r, 8), F32(r, 12), F32(r, 16), F32(r, 20));
                var (x, y, w, h) = Rects(r, 24, 1, compressed)[0];
                Image(image, I32(r, 4), src, (x, y), (x + w, y), (x, y + h));
                break;
            }
            case 0x401B when r.Length >= 28: // DrawImagePoints
            {
                if (_objects.GetValueOrDefault(id) is not PlusImage image) break;
                var src = (F32(r, 8), F32(r, 12), F32(r, 16), F32(r, 20));
                var points = Points(r, 28, Math.Min(I32(r, 24), 3), compressed, (flags & 0x0800) != 0);
                if (points.Count == 3) Image(image, I32(r, 4), src, points[0], points[1], points[2]);
                break;
            }
            case 0x401C when r.Length >= 28: // DrawString
            {
                if (_objects.GetValueOrDefault(id) is not PlusFont font) break;
                int length = I32(r, 8);
                if (length < 0 || 28 + 2L * length > r.Length) break;
                var format = _objects.GetValueOrDefault(I32(r, 4)) as PlusStringFormat;
                DrawString(Encoding.Unicode.GetString(r.Slice(28, 2 * length)), font, format, BrushOf(flags, I32(r, 0)), (F32(r, 12), F32(r, 16), F32(r, 20), F32(r, 24)));
                break;
            }
            case 0x4036 when r.Length >= 16: // DrawDriverString
            {
                if (_objects.GetValueOrDefault(id) is not PlusFont font) break;
                int options = I32(r, 4), matrixPresent = I32(r, 8), count = I32(r, 12);
                if (count < 0 || 16 + 10L * count > r.Length) break;
                string text = Encoding.Unicode.GetString(r.Slice(16, 2 * count));
                var positions = new List<(double X, double Y)>(count);
                for (int k = 0; k < count; k++) positions.Add((F32(r, 16 + 2 * count + 8 * k), F32(r, 20 + 2 * count + 8 * k)));
                int matrixAt = 16 + 10 * count;
                Affine? matrix = matrixPresent != 0 && matrixAt + 24 <= r.Length ? Xform(r, matrixAt) : null;
                DrawGlyphs(text, positions, (options & 4) != 0, matrix, font, BrushOf(flags, I32(r, 0)));
                break;
            }
            case 0x4025 or 0x4028 when r.Length >= 4: // Save, BeginContainerNoParams
                _saved[I32(r, 0)] = _state.Copy();
                break;
            case 0x4027 when r.Length >= 36: // BeginContainer
            {
                _saved[I32(r, 32)] = _state.Copy();
                var (dx, dy, dw, dh) = (F32(r, 0), F32(r, 4), F32(r, 8), F32(r, 12));
                var (sx, sy, sw, sh) = (F32(r, 16), F32(r, 20), F32(r, 24), F32(r, 28));
                // The source rectangle (in the flags' unit) maps onto the destination (world units of the outer state).
                double unit = ToWorld(1, flags >> 8 & 0xFF);
                double scaleX = sw == 0 ? 1 : dw / sw * unit, scaleY = sh == 0 ? 1 : dh / sh * unit;
                var container = new Affine(scaleX, 0, 0, scaleY, dx - sx * scaleX, dy - sy * scaleY);
                _state.World = container.Then(_state.World);
                break;
            }
            case 0x4026 or 0x4029 when r.Length >= 4: // Restore, EndContainer
                if (_saved.Remove(I32(r, 0), out var saved)) _state = saved;
                break;
            case 0x402A when r.Length >= 24: // SetWorldTransform
                _state.World = Xform(r, 0);
                break;
            case 0x402B: // ResetWorldTransform
                _state.World = Affine.Identity;
                break;
            case 0x402C when r.Length >= 24: // MultiplyWorldTransform
                Multiply(Xform(r, 0), flags);
                break;
            case 0x402D when r.Length >= 8: // TranslateWorldTransform
                Multiply(new Affine(1, 0, 0, 1, F32(r, 0), F32(r, 4)), flags);
                break;
            case 0x402E when r.Length >= 8: // ScaleWorldTransform
                Multiply(new Affine(F32(r, 0), 0, 0, F32(r, 4), 0, 0), flags);
                break;
            case 0x402F when r.Length >= 4: // RotateWorldTransform
            {
                double a = F32(r, 0) * Math.PI / 180, c = Math.Cos(a), s = Math.Sin(a);
                Multiply(new Affine(c, s, -s, c, 0, 0), flags);
                break;
            }
            case 0x4030 when r.Length >= 4: // SetPageTransform
                _state.PageUnit = flags & 0xFF;
                _state.PageScale = F32(r, 0) is > 0 and var scale ? scale : 1;
                break;
            case 0x4031: // ResetClip
                _state.Clip = [];
                break;
            case 0x4032 when r.Length >= 16: // SetClipRect
            {
                var (x, y, w, h) = (F32(r, 0), F32(r, 4), F32(r, 8), F32(r, 12));
                Clip(new PlusRegion([[(x, y), (x + w, y), (x + w, y + h), (x, y + h)]]), flags >> 8 & 0xF);
                break;
            }
            case 0x4033: // SetClipPath
                if (_objects.GetValueOrDefault(id) is List<PathPoint> clipPath) Clip(new PlusRegion([ClipGroups.Flatten(clipPath)]), flags >> 8 & 0xF);
                break;
            case 0x4034: // SetClipRegion
                if (_objects.GetValueOrDefault(id) is PlusRegion clipRegion) Clip(clipRegion, flags >> 8 & 0xF);
                break;
        }
    }

    /// <summary>Prepends the transform to the world transform, or appends it when the flags' post-multiply bit is set.</summary>
    private void Multiply(Affine m, int flags) => _state.World = (flags & 0x2000) != 0 ? _state.World.Then(m) : m.Then(_state.World);

    // ---- Objects ----

    private void ObjectRecord(int flags, ReadOnlySpan<byte> r)
    {
        int id = flags & 0xFF, type = flags >> 8 & 0x7F;
        byte[] data;
        // A large object is split over several records, each of which but the last has the continue bit and the total size.
        if ((flags & 0x8000) != 0)
        {
            if (r.Length < 4) return;
            if (!_partial.TryGetValue(id, out var part)) _partial[id] = part = (I32(r, 0), new MemoryStream());
            part.Data.Write(r[4..]);
            if (part.Data.Length < part.Total) return;
            _partial.Remove(id);
            data = part.Data.ToArray();
        }
        else if (_partial.Remove(id, out var last))
        {
            last.Data.Write(r);
            data = last.Data.ToArray();
        }
        else data = r.ToArray();

        object? o = type switch
        {
            1 => ReadBrush(data),
            2 => ReadPen(data),
            3 => ReadPath(data, 0, out _),
            4 => ReadRegion(data),
            5 => ReadImage(data),
            6 => ReadFont(data),
            7 when data.Length >= 20 => new PlusStringFormat(I32(data, 4), I32(data, 12), I32(data, 16)),
            _ => null,
        };
        if (o is null) _objects.Remove(id);
        else _objects[id] = o;
    }

    private static PlusBrush? ReadBrush(ReadOnlySpan<byte> d)
    {
        if (d.Length < 12) return null;
        switch (I32(d, 4))
        {
            case 0: return ReadColor(d, 8);
            case 1 when d.Length >= 20: return Mix(ReadColor(d, 12), ReadColor(d, 16)); // hatch: fore and back colors
            case 3 when d.Length >= 32: return ReadPathGradient(d) ?? Mix(ReadColor(d, 16), ReadColor(d, 16));
            case 4 when d.Length >= 48: return ReadLinearGradient(d) ?? Mix(ReadColor(d, 32), ReadColor(d, 36));
            default: return new PlusBrush(new Color(0x80, 0x80, 0x80), 1);
        }
    }

    private static PlusBrush? ReadLinearGradient(ReadOnlySpan<byte> d)
    {
        int flags = I32(d, 8), at = 48;
        var rect = ((double)F32(d, 16), (double)F32(d, 20), (double)F32(d, 24), (double)F32(d, 28));
        PlusBrush start = ReadColor(d, 32), end = ReadColor(d, 36);
        var transform = Affine.Identity;
        if ((flags & 0x02) != 0)
        {
            if (at + 24 > d.Length) return null;
            transform = new Affine(F32(d, at), F32(d, at + 4), F32(d, at + 8), F32(d, at + 12), F32(d, at + 16), F32(d, at + 20));
            at += 24;
        }
        var stops = GradientStops(d, at, flags, start, end, out double alpha);
        if (stops is null || rect.Item3 == 0) return null;
        return Average(stops, alpha) with { Gradient = new PlusGradient(false, rect, (0, 0), transform, stops) { Wrap = I32(d, 12) } };
    }

    private static PlusBrush? ReadPathGradient(ReadOnlySpan<byte> d)
    {
        int flags = I32(d, 8), count = I32(d, 28), at = 32 + 4 * count;
        if (count < 1 || at + 4 > d.Length) return null;
        PlusBrush center = ReadColor(d, 16), surround = ReadColor(d, 32);
        var centerPoint = ((double)F32(d, 20), (double)F32(d, 24));
        List<(double X, double Y)> boundary;
        if ((flags & 0x01) != 0)
        {
            int size = I32(d, at);
            if (size < 0 || at + 4 + size > d.Length) return null;
            boundary = ReadPath(d, at + 4, out _) is { } path ? ClipGroups.Flatten(path).Where(p => !double.IsNaN(p.X)).ToList() : [];
            at += 4 + size;
        }
        else
        {
            int points = I32(d, at);
            if (points < 0 || at + 4 + 8 * points > d.Length) return null;
            boundary = Points(d, at + 4, points, false, false);
            at += 4 + 8 * points;
        }
        if (boundary.Count < 2) return null;
        var transform = Affine.Identity;
        if ((flags & 0x02) != 0)
        {
            if (at + 24 > d.Length) return null;
            transform = new Affine(F32(d, at), F32(d, at + 4), F32(d, at + 8), F32(d, at + 12), F32(d, at + 16), F32(d, at + 20));
            at += 24;
        }
        // The blend runs from the boundary (position 0) to the center (position 1); stops here run from the center out.
        var stops = GradientStops(d, at, flags, surround, center, out double alpha);
        if (stops is null) return null;
        stops = stops.Select(s => new GradientStop(1 - s.Position, s.Color)).Reverse().ToList();
        double left = boundary.Min(p => p.X), top = boundary.Min(p => p.Y);
        var rect = (left, top, boundary.Max(p => p.X) - left, boundary.Max(p => p.Y) - top);
        bool rectangular = boundary.Distinct().Count() == 4 && boundary.All(p => (p.X == left || p.X == left + rect.Item3) && (p.Y == top || p.Y == top + rect.Item4))
            && transform.B == 0 && transform.C == 0;
        return Average(stops, alpha) with { Gradient = new PlusGradient(true, rect, centerPoint, transform, stops) { Rectangular = rectangular } };
    }

    /// <summary>
    /// A gradient's stops from its preset colors (flag 4) or blend factors (flag 8) after <paramref name="at"/>, else from
    /// <paramref name="start"/> to <paramref name="end"/>; <paramref name="alpha"/> is their mean opacity.
    /// </summary>
    private static List<GradientStop>? GradientStops(ReadOnlySpan<byte> d, int at, int flags, PlusBrush start, PlusBrush end, out double alpha)
    {
        alpha = (start.Alpha + end.Alpha) / 2;
        if ((flags & 0x0C) == 0) return [new(0, start.Color), new(1, end.Color)];
        if (at + 4 > d.Length) return null;
        int count = I32(d, at);
        if (count < 1 || count > (d.Length - at - 4) / 8) return null;
        var stops = new List<GradientStop>(count);
        double sum = 0;
        for (int k = 0; k < count; k++)
        {
            double position = Math.Clamp(F32(d, at + 4 + 4 * k), 0, 1);
            if ((flags & 0x04) != 0)
            {
                var color = ReadColor(d, at + 4 + 4 * count + 4 * k);
                stops.Add(new(position, color.Color));
                sum += color.Alpha;
            }
            else
            {
                double factor = Math.Clamp(F32(d, at + 4 + 4 * count + 4 * k), 0, 1);
                stops.Add(new(position, start.Color.Blend(end.Color, factor)));
                sum += start.Alpha + (end.Alpha - start.Alpha) * factor;
            }
        }
        alpha = sum / count;
        return stops.OrderBy(s => s.Position).ToList();
    }

    private static PlusBrush Average(List<GradientStop> stops, double alpha) =>
        new(new Color((byte)stops.Average(s => s.Color.R), (byte)stops.Average(s => s.Color.G), (byte)stops.Average(s => s.Color.B)), alpha);

    private static PlusBrush Mix(PlusBrush a, PlusBrush b) =>
        new(new Color((byte)((a.Color.R + b.Color.R) / 2), (byte)((a.Color.G + b.Color.G) / 2), (byte)((a.Color.B + b.Color.B) / 2)), (a.Alpha + b.Alpha) / 2);

    private static PlusPen? ReadPen(ReadOnlySpan<byte> d)
    {
        if (d.Length < 20) return null;
        int flags = I32(d, 8), unit = I32(d, 12);
        double width = F32(d, 16);
        int at = 20, lineStyle = 0;
        // The optional data, in the order of its flag bits.
        if ((flags & 0x1) != 0) at += 24;
        if ((flags & 0x2) != 0) at += 4;
        if ((flags & 0x4) != 0) at += 4;
        if ((flags & 0x8) != 0) at += 4;
        if ((flags & 0x10) != 0) at += 4;
        if ((flags & 0x20) != 0) { if (at + 4 > d.Length) return null; lineStyle = I32(d, at); at += 4; }
        if ((flags & 0x40) != 0) at += 4;
        if ((flags & 0x80) != 0) at += 4;
        if ((flags & 0x100) != 0) { if (at + 4 > d.Length) return null; at += 4 + 4 * I32(d, at); }
        if ((flags & 0x200) != 0) at += 4;
        if ((flags & 0x400) != 0) { if (at + 4 > d.Length) return null; at += 4 + 4 * I32(d, at); }
        if ((flags & 0x800) != 0) { if (at + 4 > d.Length) return null; at += 4 + I32(d, at); }
        if ((flags & 0x1000) != 0) { if (at + 4 > d.Length) return null; at += 4 + I32(d, at); }
        if (at < 0 || at > d.Length) return null;
        var brush = ReadBrush(d[at..]) ?? new PlusBrush(Color.Black, 1);
        return new PlusPen(width, unit, (flags & 0x20) != 0 ? lineStyle : 0, brush);
    }

    /// <summary>A path object's points (world coordinates) and figure types; <paramref name="size"/> is the bytes it takes.</summary>
    private static List<PathPoint>? ReadPath(ReadOnlySpan<byte> d, int at, out int size)
    {
        size = 0;
        if (at + 12 > d.Length) return null;
        int count = I32(d, at + 4), flags = I32(d, at + 8);
        if (count < 0 || (flags & 0x0800) != 0) return null; // relative (run-length) points are not played
        bool small = (flags & 0x4000) != 0;
        int pointsAt = at + 12, typesAt = pointsAt + count * (small ? 4 : 8);
        if (typesAt + count > d.Length) return null;
        size = 12 + count * (small ? 4 : 8) + (count + 3) / 4 * 4;
        var path = new List<PathPoint>(count + 4);
        for (int k = 0; k < count; k++)
        {
            double x = small ? I16(d, pointsAt + 4 * k) : F32(d, pointsAt + 8 * k);
            double y = small ? I16(d, pointsAt + 4 * k + 2) : F32(d, pointsAt + 8 * k + 4);
            int t = d[typesAt + k];
            var kind = (t & 7) switch { 0 => PathPointKind.Move, 3 => PathPointKind.Curve, _ => PathPointKind.Line };
            if (k == 0) kind = PathPointKind.Move;
            path.Add(new PathPoint(x, y, kind));
            if ((t & 0x80) != 0) path.Add(new PathPoint(0, 0, PathPointKind.Close));
        }
        return path;
    }

    private static PlusRegion? ReadRegion(ReadOnlySpan<byte> d)
    {
        int at = 8;
        return d.Length >= 8 ? RegionNode(d, ref at) : null;
    }

    /// <summary>
    /// A region node: rectangles, paths and infinite regions, intersections, and unions, differences and exclusive-ors of single outlines
    /// (played as even-odd rings, right when the parts nest or are apart). Null for what cannot be played.
    /// </summary>
    private static PlusRegion? RegionNode(ReadOnlySpan<byte> d, ref int at)
    {
        if (at + 4 > d.Length) return null;
        int type = I32(d, at);
        at += 4;
        switch (type)
        {
            case 0x10000000 when at + 16 <= d.Length:
            {
                var (x, y, w, h) = (F32(d, at), F32(d, at + 4), F32(d, at + 8), F32(d, at + 12));
                at += 16;
                return new PlusRegion([[(x, y), (x + w, y), (x + w, y + h), (x, y + h)]]);
            }
            case 0x10000001 when at + 4 <= d.Length:
            {
                int length = I32(d, at);
                var path = ReadPath(d, at + 4, out _);
                at += 4 + length;
                return path is null ? null : new PlusRegion([ClipGroups.Flatten(path)]);
            }
            case 0x10000002:
                return new PlusRegion([[]], Empty: true);
            case 0x10000003:
                return new PlusRegion([]);
            case >= 1 and <= 4 or 0:
            {
                var left = RegionNode(d, ref at);
                var right = RegionNode(d, ref at);
                if (left is null || right is null) return null;
                if (type == 0) return left.Empty || right.Empty ? new PlusRegion([[]], true) : new PlusRegion([.. left.Outlines, .. right.Outlines]);
                if (left.Empty) return type is 1 or 2 ? right : left;
                if (right.Empty) return left;
                if (left.Outlines.Count != 1 || right.Outlines.Count != 1) return null;
                return new PlusRegion([[.. left.Outlines[0], (double.NaN, double.NaN), .. right.Outlines[0]]]);
            }
            default:
                return null;
        }
    }

    private static PlusImage? ReadImage(byte[] d)
    {
        if (d.Length < 8) return null;
        int type = I32(d, 4);
        if (type == 2 && d.Length >= 16)
        {
            // A metafile (WMF, placeable WMF or EMF): drawn as a nested picture.
            int size = Math.Min(I32(d, 12), d.Length - 16);
            if (size <= 0) return null;
            var bytes = d.AsSpan(16, size).ToArray();
            return new PlusImage(new ImageData(bytes, "EMF+ picture"), 0, 0);
        }
        if (type != 1 || d.Length < 28) return null;
        int width = I32(d, 8), height = I32(d, 12), stride = I32(d, 16), format = I32(d, 20), kind = I32(d, 24);
        if (kind == 1) return new PlusImage(new ImageData(d.AsSpan(28).ToArray(), "EMF+ picture"), width, height);
        if (width <= 0 || height <= 0 || stride <= 0 || 28L + (long)stride * height > d.Length) return null;
        int bpp = format switch { 0x0026200A or 0x000E200B or 0x00022009 => 32, 0x00021808 => 24, _ => 0 };
        if (bpp == 0) return new PlusImage(null, width, height);
        // Rows run top to bottom: a BMP with a negative height.
        var info = new byte[56];
        BinaryPrimitives.WriteInt32LittleEndian(info, 40);
        BinaryPrimitives.WriteInt32LittleEndian(info.AsSpan(4), width);
        BinaryPrimitives.WriteInt32LittleEndian(info.AsSpan(8), -height);
        BinaryPrimitives.WriteInt16LittleEndian(info.AsSpan(12), 1);
        BinaryPrimitives.WriteInt16LittleEndian(info.AsSpan(14), (short)bpp);
        int bmpStride = (width * bpp + 31) / 32 * 4;
        var bits = new byte[bmpStride * height];
        for (int row = 0; row < height; row++) d.AsSpan(28 + row * stride, Math.Min(stride, bmpStride)).CopyTo(bits.AsSpan(row * bmpStride));
        ImageData? image;
        if (format == 0x000E200B) image = GdiPlayer.AlphaDib(info.AsSpan(0, 40), bits, 255);
        else if (format == 0x0026200A)
        {
            BinaryPrimitives.WriteInt32LittleEndian(info, 56);
            BinaryPrimitives.WriteInt32LittleEndian(info.AsSpan(16), 3);
            BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(40), 0x00FF0000);
            BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(44), 0x0000FF00);
            BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(48), 0x000000FF);
            BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(52), 0xFF000000);
            image = GdiPlayer.Dib(info, bits);
        }
        else image = GdiPlayer.Dib(info.AsSpan(0, 40), bits);
        return new PlusImage(image, width, height);
    }

    private static PlusFont? ReadFont(ReadOnlySpan<byte> d)
    {
        if (d.Length < 24) return null;
        int length = I32(d, 20);
        if (length < 0 || 24 + 2L * length > d.Length) return null;
        return new PlusFont(F32(d, 4), I32(d, 8), I32(d, 12), Encoding.Unicode.GetString(d.Slice(24, 2 * length)));
    }

    private PlusBrush? BrushOf(int flags, int value) =>
        (flags & 0x8000) != 0 ? ReadColor(BitConverter.GetBytes(value), 0) : _objects.GetValueOrDefault(value & 0xFF) as PlusBrush;

    private PlusPen? PenOf(int id) => _objects.GetValueOrDefault(id & 0xFF) as PlusPen;

    // ---- Drawing ----

    private static PathPoint P(double x, double y, PathPointKind kind = PathPointKind.Line) => new(x, y, kind);
    private static readonly PathPoint Close = new(0, 0, PathPointKind.Close);

    private static List<PathPoint> RectPath(double x, double y, double w, double h) =>
        [P(x, y, PathPointKind.Move), P(x + w, y), P(x + w, y + h), P(x, y + h), Close];

    private static List<PathPoint> Polyline(List<(double X, double Y)> points, bool closed)
    {
        var path = points.Select((p, k) => P(p.X, p.Y, k == 0 ? PathPointKind.Move : PathPointKind.Line)).ToList();
        if (closed && path.Count > 0) path.Add(Close);
        return path;
    }

    private static List<PathPoint> Ellipse(double x, double y, double w, double h) => Arc(x, y, w, h, 0, 360, false);

    /// <summary>
    /// An elliptical arc in Bézier segments; angles are in degrees clockwise from the x axis, measured to the point on the ellipse (as GDI+
    /// does). A pie closes through the center.
    /// </summary>
    private static List<PathPoint> Arc(double x, double y, double w, double h, double start, double sweep, bool pie)
    {
        double rx = w / 2, ry = h / 2, cx = x + rx, cy = y + ry;
        sweep = Math.Clamp(sweep, -360, 360);
        // The ellipse's parameter for a geometric angle.
        double Param(double degrees)
        {
            double a = degrees * Math.PI / 180;
            double t = Math.Atan2(rx * Math.Sin(a), ry * Math.Cos(a));
            // Keep whole turns so the sweep keeps its length.
            return t + Math.Round((a - t) / (2 * Math.PI)) * 2 * Math.PI;
        }
        double t0 = Param(start), t1 = Math.Abs(sweep) >= 360 ? t0 + Math.Sign(sweep) * 2 * Math.PI : Param(start + sweep);
        var path = new List<PathPoint>();
        if (pie) path.Add(P(cx, cy, PathPointKind.Move));
        path.Add(P(cx + rx * Math.Cos(t0), cy + ry * Math.Sin(t0), pie ? PathPointKind.Line : PathPointKind.Move));
        int segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(t1 - t0) / (Math.PI / 2) - 1e-9));
        double step = (t1 - t0) / segments, k = 4.0 / 3 * Math.Tan(step / 4);
        for (int s = 0; s < segments; s++)
        {
            double a = t0 + s * step, b = a + step;
            path.Add(P(cx + rx * (Math.Cos(a) - k * Math.Sin(a)), cy + ry * (Math.Sin(a) + k * Math.Cos(a)), PathPointKind.Curve));
            path.Add(P(cx + rx * (Math.Cos(b) + k * Math.Sin(b)), cy + ry * (Math.Sin(b) - k * Math.Cos(b)), PathPointKind.Curve));
            path.Add(P(cx + rx * Math.Cos(b), cy + ry * Math.Sin(b), PathPointKind.Curve));
        }
        if (pie || Math.Abs(sweep) >= 360) path.Add(Close);
        return path;
    }

    /// <summary>A cardinal spline through the points as Bézier segments (GDI+'s tension: 0 straight, 0.5 the default).</summary>
    private static List<PathPoint> Curve(List<(double X, double Y)> points, double tension, bool closed)
    {
        var path = new List<PathPoint>();
        int n = points.Count;
        if (n < 2) return path;
        (double X, double Y) At(int k) => closed ? points[(k % n + n) % n] : points[Math.Clamp(k, 0, n - 1)];
        double t = tension / 3;
        path.Add(P(points[0].X, points[0].Y, PathPointKind.Move));
        int segments = closed ? n : n - 1;
        for (int k = 0; k < segments; k++)
        {
            var (p0, p1, p2, p3) = (At(k - 1), At(k), At(k + 1), At(k + 2));
            path.Add(P(p1.X + t * (p2.X - p0.X), p1.Y + t * (p2.Y - p0.Y), PathPointKind.Curve));
            path.Add(P(p2.X - t * (p3.X - p1.X), p2.Y - t * (p3.Y - p1.Y), PathPointKind.Curve));
            path.Add(P(p2.X, p2.Y, PathPointKind.Curve));
        }
        if (closed) path.Add(Close);
        return path;
    }

    private List<PathPoint> ToBox(IReadOnlyList<PathPoint> world, Affine m) =>
        world.Select(p =>
        {
            if (p.Kind == PathPointKind.Close) return p;
            var (x, y) = m.Apply(p.X, p.Y);
            return p with { X = x, Y = y };
        }).ToList();

    /// <summary>Adds an item, inside a group of the brush's or pen's opacity when it is translucent.</summary>
    private void AddWithAlpha(LayoutItem item, double alpha)
    {
        if (alpha <= 0) return;
        if (alpha >= 1) { Add(item); return; }
        Add(new PushItem(null, Opacity: alpha));
        Add(item);
        Add(new PopItem());
    }

    /// <summary>Fills a path (world coordinates, or with <paramref name="box"/> already in the output box) with the brush.</summary>
    private void Fill(List<PathPoint> path, PlusBrush? brush, bool evenOdd = true, bool box = false)
    {
        if (brush is null || path.Count < 3) return;
        var points = box ? path : ToBox(path, Transform);
        if (brush.Gradient is { } gradient && GradientFill(points, gradient) is { } fill)
        {
            // The gradient is painted in the path's bounds, clipped to the path (nonzero fills are clipped even-odd, as the clip groups do).
            var saved = _state.Clip;
            _state.Clip = [.. saved, ClipGroups.Flatten(points)];
            AddWithAlpha(fill, brush.Alpha);
            _state.Clip = saved;
            return;
        }
        AddWithAlpha(new PathItem(points, brush.Color, null, 0.05, LineStyle.Solid, evenOdd), brush.Alpha);
    }

    /// <summary>A gradient item over the bounds of <paramref name="points"/> (output box coordinates) that paints the gradient brush.</summary>
    private GradientItem? GradientFill(List<PathPoint> points, PlusGradient gradient)
    {
        var positions = points.Where(p => p.Kind != PathPointKind.Close).ToList();
        if (positions.Count == 0) return null;
        double left = positions.Min(p => p.X), top = positions.Min(p => p.Y);
        double width = positions.Max(p => p.X) - left, height = positions.Max(p => p.Y) - top;
        if (width <= 0 || height <= 0) return null;
        var m = gradient.Transform.Then(Transform);
        var (gx, gy, gw, gh) = gradient.Rect;
        if (gradient.Radial)
        {
            // The blend reaches the stop at 1 at the bounds of the boundary; the focus is the center point.
            var corners = new[] { m.Apply(gx, gy), m.Apply(gx + gw, gy), m.Apply(gx, gy + gh), m.Apply(gx + gw, gy + gh) };
            double rl = corners.Min(c => c.X), rt = corners.Min(c => c.Y), rw = corners.Max(c => c.X) - rl, rh = corners.Max(c => c.Y) - rt;
            if (rw <= 0 || rh <= 0) return null;
            var (cx, cy) = m.Apply(gradient.Center.X, gradient.Center.Y);
            return new GradientItem(rl, rt, rw, rh, 0, gradient.Stops)
                { Radial = true, Rectangular = gradient.Rectangular && Math.Abs(Transform.B) < 1e-9 && Math.Abs(Transform.C) < 1e-9, FocusX = Math.Clamp((cx - rl) / rw, 0, 1), FocusY = Math.Clamp((cy - rt) / rh, 0, 1) };
        }
        // The blend position at a box point is the brush-space x of it, from the rectangle's left (0) to its right edge (1): the gradient
        // runs along the first row of the inverse transform.
        double det = m.A * m.D - m.B * m.C;
        if (Math.Abs(det) < 1e-12) return null;
        double ia = m.D / det, ic = -m.C / det, ie = (m.C * m.F - m.D * m.E) / det;
        double rate = Math.Sqrt(ia * ia + ic * ic) / gw; // blend positions per unit of length along the gradient
        if (rate == 0) return null;
        double degree = Math.Atan2(ic * Math.Sign(gw), ia * Math.Sign(gw)) * 180 / Math.PI;
        double rad = degree * Math.PI / 180, half = (Math.Abs(width * Math.Cos(rad)) + Math.Abs(height * Math.Sin(rad))) / 2;
        double centerAt = (ia * (left + width / 2) + ic * (top + height / 2) + ie - gx) / gw;
        rate = Math.Abs(rate);
        // A blend position t sits at (t - centerAt) / rate from the box center; positions outside the rectangle keep the end colors.
        double ToBoxStop(double t) => ((t - centerAt) / rate + half) / (2 * half);
        // Tiling repeats the stops in every period of the rectangle that reaches the box; flipped tiling reverses every other copy.
        var stops = new List<GradientStop>();
        if (gradient.Wrap is >= 0 and < 4)
        {
            double from = centerAt - half * rate, to = centerAt + half * rate;
            if (to - from > 200) return null; // far too many copies to list
            for (int period = (int)Math.Floor(from + 1e-9); period < to - 1e-9; period++)
            {
                bool flip = gradient.Wrap != 0 && (period & 1) != 0 && gradient.Wrap != 2;
                var copy = gradient.Stops.Select(s => new GradientStop(period + (flip ? 1 - s.Position : s.Position), s.Color));
                stops.AddRange(flip ? copy.Reverse() : copy);
            }
        }
        else stops.AddRange(gradient.Stops);
        return new GradientItem(left, top, width, height, degree, ClampStops(stops.Select(s => s with { Position = ToBoxStop(s.Position) }).ToList()));
    }

    /// <summary>Stops cut to positions 0 to 1, with the colors they blend to at the cut ends.</summary>
    private static List<GradientStop> ClampStops(List<GradientStop> stops)
    {
        Color At(double t)
        {
            if (t <= stops[0].Position) return stops[0].Color;
            for (int k = 1; k < stops.Count; k++)
                if (t <= stops[k].Position)
                    return stops[k - 1].Color.Blend(stops[k].Color, (t - stops[k - 1].Position) / Math.Max(stops[k].Position - stops[k - 1].Position, 1e-12));
            return stops[^1].Color;
        }
        var result = new List<GradientStop> { new(0, At(0)) };
        result.AddRange(stops.Where(s => s.Position > 1e-9 && s.Position < 1 - 1e-9));
        result.Add(new(1, At(1)));
        return result;
    }

    private void Stroke(List<PathPoint> path, PlusPen pen)
    {
        if (path.Count < 2) return;
        var m = Transform;
        // A zero width is one device pixel, whatever the transform.
        double width = pen.Width <= 0 ? DeviceToBox.Scale : ToWorld(pen.Width, pen.Unit) * m.Scale;
        var style = pen.LineStyle switch { 1 or 3 or 4 => LineStyle.Dashed, 2 => LineStyle.Dotted, _ => LineStyle.Solid };
        AddWithAlpha(new PathItem(ToBox(path, m), null, pen.Brush.Color, Math.Max(width, 0.05), style), pen.Brush.Alpha);
    }

    private void FillRegion(PlusRegion region, PlusBrush? brush)
    {
        if (brush is null || region.Empty) return;
        var m = Transform;
        var outlines = region.Outlines.Select(o => Box(o, m)).ToList();
        if (outlines.Count == 0)
        {
            double h = gdi.BoxHeight;
            outlines.Add([(0, 0), (Metafiles.UnitBox, 0), (Metafiles.UnitBox, h), (0, h)]);
        }
        var saved = _state.Clip;
        _state.Clip = [.. saved, .. outlines.Take(outlines.Count - 1)];
        var path = new List<PathPoint>();
        foreach (var p in outlines[^1])
        {
            if (double.IsNaN(p.X)) { path.Add(Close); continue; }
            path.Add(P(p.X, p.Y, path.Count == 0 || path[^1].Kind == PathPointKind.Close ? PathPointKind.Move : PathPointKind.Line));
        }
        path.Add(Close);
        Fill(path, brush, true, box: true);
        _state.Clip = saved;
    }

    private static List<(double X, double Y)> Box(List<(double X, double Y)> outline, Affine m) =>
        outline.Select(p => double.IsNaN(p.X) ? p : m.Apply(p.X, p.Y)).ToList();

    /// <summary>Combines a region (world coordinates) with the clip: replace 0, intersect 1, exclude 4; union and the others leave it.</summary>
    private void Clip(PlusRegion region, int mode)
    {
        var m = Transform;
        var outlines = region.Outlines.Select(o => (IReadOnlyList<(double X, double Y)>)Box(o, m)).ToArray();
        switch (mode)
        {
            case 0: _state.Clip = outlines; break;
            case 1: _state.Clip = [.. _state.Clip, .. outlines]; break;
            case 4 when outlines.Length == 1:
            {
                // Outside the region: its rings inside a ring around everything clip even-odd.
                double h = gdi.BoxHeight, big = 1e5;
                List<(double X, double Y)> outside = [(-big, -big), (big, -big), (big, h + big), (-big, h + big), (double.NaN, double.NaN), .. outlines[0]];
                _state.Clip = [.. _state.Clip, outside];
                break;
            }
        }
    }

    /// <summary>
    /// Draws an image so its source rectangle (in <paramref name="srcUnit"/>, pixels as a rule) fills the parallelogram with these corners
    /// (top-left, top-right, bottom-left; world coordinates).
    /// </summary>
    private void Image(PlusImage image, int srcUnit, (double X, double Y, double W, double H) src, (double X, double Y) p0, (double X, double Y) p1, (double X, double Y) p2)
    {
        if (image.Data is not { } data || src.W == 0 || src.H == 0) return;
        var m = Transform;
        var (a, b, c) = (m.Apply(p0.X, p0.Y), m.Apply(p1.X, p1.Y), m.Apply(p2.X, p2.Y));
        double width = Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
        double height = Math.Sqrt((c.X - a.X) * (c.X - a.X) + (c.Y - a.Y) * (c.Y - a.Y));
        if (width <= 0 || height <= 0) return;
        double angle = Math.Atan2(b.Y - a.Y, b.X - a.X);
        // A mirrored parallelogram: the second side turns the other way.
        bool flip = (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X) < 0;
        bool turned = Math.Abs(angle) > 1e-6 || flip;
        if (turned) Add(new PushItem(null, -angle * 180 / Math.PI, a.X, a.Y, false, flip));

        // The whole image, placed so the source rectangle lands on the destination; clipped to it when the source is a part.
        var (imageW, imageH) = (image.Width, image.Height);
        if ((imageW <= 0 || imageH <= 0) && data.PixelSize is { } px) (imageW, imageH) = (px.Width, px.Height);
        double unit = srcUnit is 2 or 0 || imageW <= 0 ? 1 : UnitScale(srcUnit, false) * 96 / _dpiX;
        bool whole = imageW <= 0 || imageH <= 0
            || Math.Abs(src.X) < 0.5 && Math.Abs(src.Y) < 0.5 && Math.Abs(src.W * unit - imageW) < 0.5 && Math.Abs(src.H * unit - imageH) < 0.5;
        if (whole) Add(new ImageItem(a.X, a.Y, width, height, data));
        else
        {
            double sx = width / (src.W * unit), sy = height / (src.H * unit);
            Add(new PushItem((a.X, a.Y, width, height)));
            Add(new ImageItem(a.X - src.X * unit * sx, a.Y - src.Y * unit * sy, imageW * sx, imageH * sy, data));
            Add(new PopItem());
        }
        if (turned) Add(new PopItem());
    }

    private (TrueTypeFont Font, double Size, ResolvedFont Resolved)? Font(PlusFont font, Affine m)
    {
        var resolved = fonts.Resolve(font.Family, (font.Style & 1) != 0, (font.Style & 2) != 0);
        double unitY = Math.Sqrt(m.C * m.C + m.D * m.D);
        double size = ToWorld(font.EmSize, font.Unit) * unitY;
        return size > 0 ? (resolved.Font, size, resolved) : null;
    }

    /// <summary>
    /// DrawString: the text laid out in the rectangle (wrapped at its width unless the format says not to, aligned by the format), or from
    /// the rectangle's corner when it has no size.
    /// </summary>
    private void DrawString(string text, PlusFont font, PlusStringFormat? format, PlusBrush? brush, (double X, double Y, double W, double H) rect)
    {
        if (brush is null || text.Length == 0) return;
        var m = Transform;
        if (Font(font, m) is not { } resolvedFont) return;
        var (ttf, size, resolved) = resolvedFont;
        double unitX = Math.Sqrt(m.A * m.A + m.B * m.B), unitY = Math.Sqrt(m.C * m.C + m.D * m.D);
        double boxW = rect.W * unitX, boxH = rect.H * unitY;
        var metrics = FontMetrics.Of(ttf, size);
        bool wrap = boxW > 0 && (format is null || (format.Flags & 0x1000) == 0);

        var lines = new List<string>();
        foreach (var paragraph in text.Replace("\r\n", "\n").Split('\n', '\r'))
        {
            if (!wrap) { lines.Add(paragraph); continue; }
            var line = new StringBuilder();
            foreach (var word in paragraph.Split(' '))
            {
                string candidate = line.Length == 0 ? word : line + " " + word;
                if (line.Length > 0 && ttf.MeasureText(candidate, size) > boxW) { lines.Add(line.ToString()); line.Clear().Append(word); }
                else line.Clear().Append(candidate);
            }
            lines.Add(line.ToString());
        }

        int align = format?.Alignment ?? 0, lineAlign = format?.LineAlignment ?? 0;
        double lineHeight = metrics.Height, total = lineHeight * lines.Count;
        double top = boxH <= 0 ? 0 : lineAlign switch { 1 => (boxH - total) / 2, 2 => boxH - total, _ => 0 };
        var (ox, oy) = m.Apply(rect.X, rect.Y);
        double rotation = -Math.Atan2(m.B, m.A) * 180 / Math.PI;
        bool rotated = Math.Abs(rotation % 360) > 0.01;
        if (rotated) Add(new PushItem(null, rotation, ox, oy));
        var color = brush.Color;
        if (brush.Alpha < 1) Add(new PushItem(null, Opacity: brush.Alpha));
        for (int k = 0; k < lines.Count; k++)
        {
            double width = ttf.MeasureText(lines[k], size);
            double area = boxW > 0 ? boxW : 0;
            double x = align switch { 1 => (area - width) / 2, 2 => area - width, _ => 0 };
            double baseline = oy + top + k * lineHeight + metrics.Ascent;
            if (lines[k].Length > 0) Add(new TextItem(ox + x, baseline, ttf, size, lines[k], color, resolved.SyntheticBold, resolved.SyntheticItalic));
            Decorate(font, ttf, size, ox + x, baseline, width, color);
        }
        if (brush.Alpha < 1) Add(new PopItem());
        if (rotated) Add(new PopItem());
    }

    private void Decorate(PlusFont font, TrueTypeFont ttf, double size, double x, double baseline, double width, Color color)
    {
        double thickness = Math.Max(size * ttf.UnderlineThickness / ttf.UnitsPerEm, size / 20);
        if ((font.Style & 4) != 0)
        {
            double y = baseline - size * ttf.UnderlinePosition / ttf.UnitsPerEm;
            Add(new LineItem(x, y, x + width, y, thickness, color));
        }
        if ((font.Style & 8) != 0)
        {
            double y = baseline - size * (ttf.StrikeoutPosition != 0 ? ttf.StrikeoutPosition : ttf.UnitsPerEm / 4) / ttf.UnitsPerEm;
            Add(new LineItem(x, y, x + width, y, thickness, color));
        }
    }

    /// <summary>
    /// DrawDriverString: characters at their own baseline positions (world coordinates), or from the first position with their natural
    /// advances (RealizedAdvance). Glyph indices (without the CmapLookup option) are taken as characters, as LibreOffice does.
    /// </summary>
    private void DrawGlyphs(string text, List<(double X, double Y)> positions, bool realized, Affine? matrix, PlusFont font, PlusBrush? brush)
    {
        if (brush is null || text.Length == 0 || positions.Count == 0) return;
        var m = matrix is { } g ? g.Then(Transform) : Transform;
        if (Font(font, Transform) is not { } resolvedFont) return;
        var (ttf, size, resolved) = resolvedFont;
        double rotation = -Math.Atan2(m.B, m.A) * 180 / Math.PI;
        bool rotated = Math.Abs(rotation % 360) > 0.01;
        var (ox, oy) = m.Apply(positions[0].X, positions[0].Y);
        if (rotated) Add(new PushItem(null, rotation, ox, oy));
        if (brush.Alpha < 1) Add(new PushItem(null, Opacity: brush.Alpha));
        double width = ttf.MeasureText(text, size);
        if (realized) Add(new TextItem(ox, oy, ttf, size, text, brush.Color, resolved.SyntheticBold, resolved.SyntheticItalic));
        else
        {
            // Unrotated positions relative to the first, so the rotation group turns them into place.
            double cos = Math.Cos(rotation * Math.PI / 180), sin = Math.Sin(rotation * Math.PI / 180);
            for (int k = 0; k < text.Length && k < positions.Count; k++)
            {
                if (char.IsWhiteSpace(text[k])) continue;
                var (px, py) = m.Apply(positions[k].X, positions[k].Y);
                double dx = px - ox, dy = py - oy;
                double ux = rotated ? dx * cos - dy * sin : dx, uy = rotated ? dx * sin + dy * cos : dy;
                Add(new TextItem(ox + ux, oy + uy, ttf, size, text[k].ToString(), brush.Color, resolved.SyntheticBold, resolved.SyntheticItalic));
            }
            if (positions.Count >= text.Length && text.Length > 0)
            {
                var (lx, ly) = m.Apply(positions[text.Length - 1].X, positions[text.Length - 1].Y);
                width = Math.Sqrt((lx - ox) * (lx - ox) + (ly - oy) * (ly - oy)) + ttf.MeasureText(text[^1].ToString(), size);
            }
        }
        Decorate(font, ttf, size, ox, oy, width, brush.Color);
        if (brush.Alpha < 1) Add(new PopItem());
        if (rotated) Add(new PopItem());
    }

    // ---- Reading ----

    /// <summary>An ARGB color (stored blue, green, red, alpha).</summary>
    private static PlusBrush ReadColor(ReadOnlySpan<byte> d, int o) => new(new Color(d[o + 2], d[o + 1], d[o]), d[o + 3] / 255.0);

    private static List<(double X, double Y, double W, double H)> Rects(ReadOnlySpan<byte> r, int at, int count, bool small)
    {
        int step = small ? 8 : 16;
        if (count < 0 || at + (long)count * step > r.Length) throw new InvalidDataException("A rectangle list runs past its EMF+ record.");
        var rects = new List<(double, double, double, double)>(count);
        for (int k = 0; k < count; k++)
        {
            int o = at + k * step;
            rects.Add(small ? (I16(r, o), I16(r, o + 2), I16(r, o + 4), I16(r, o + 6)) : (F32(r, o), F32(r, o + 4), F32(r, o + 8), F32(r, o + 12)));
        }
        return rects;
    }

    /// <summary>Points as floats, 16-bit integers, or (relative) each a run-length offset from the one before.</summary>
    private static List<(double X, double Y)> Points(ReadOnlySpan<byte> r, int at, int count, bool small, bool relative)
    {
        if (count < 0 || count > r.Length) throw new InvalidDataException("A point list runs past its EMF+ record.");
        var points = new List<(double X, double Y)>(count);
        if (relative)
        {
            double x = 0, y = 0;
            for (int k = 0; k < count; k++)
            {
                x += Relative(r, ref at);
                y += Relative(r, ref at);
                points.Add((x, y));
            }
            return points;
        }
        int step = small ? 4 : 8;
        if (at + (long)count * step > r.Length) throw new InvalidDataException("A point list runs past its EMF+ record.");
        for (int k = 0; k < count; k++)
        {
            int o = at + k * step;
            points.Add(small ? (I16(r, o), I16(r, o + 2)) : (F32(r, o), F32(r, o + 4)));
        }
        return points;
    }

    /// <summary>A relative coordinate: seven bits signed, or with the high bit set fifteen bits signed.</summary>
    private static int Relative(ReadOnlySpan<byte> r, ref int at)
    {
        int b = r[at++];
        if ((b & 0x80) == 0) return (b & 0x40) != 0 ? b - 0x80 : b;
        int v = (b & 0x7F) << 8 | r[at++];
        return (v & 0x4000) != 0 ? v - 0x8000 : v;
    }

    private static Affine Xform(ReadOnlySpan<byte> d, int o) => new(F32(d, o), F32(d, o + 4), F32(d, o + 8), F32(d, o + 12), F32(d, o + 16), F32(d, o + 20));
    private static int I32(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt32LittleEndian(d[o..]);
    private static int U16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadUInt16LittleEndian(d[o..]);
    private static short I16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt16LittleEndian(d[o..]);
    private static float F32(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadSingleLittleEndian(d[o..]);

    /// <summary>Whether the EMF draws through EMF+ records alone (its EMF+ header does not mark it dual).</summary>
    public static bool PlusOnly(ReadOnlySpan<byte> emf)
    {
        int pos = 0;
        while (pos + 8 <= emf.Length)
        {
            int type = I32(emf, pos), size = I32(emf, pos + 4);
            if (size < 8 || pos + size > emf.Length) return false;
            if (type == 70 && size >= 28 && emf.Slice(pos + 12, 4).SequenceEqual("EMF+"u8))
                return U16(emf, pos + 16) == 0x4001 && (U16(emf, pos + 18) & 1) == 0;
            if (type is not (1 or 70)) return false;
            pos += size;
        }
        return false;
    }
}
