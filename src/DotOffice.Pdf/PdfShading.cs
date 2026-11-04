namespace DotOffice.Pdf;

/// <summary>A color stop of a <see cref="PdfShading"/>: <see cref="Position"/> runs from 0 to 1.</summary>
public readonly record struct ShadingStop(double Position, double R, double G, double B);

/// <summary>
/// A smooth color transition painted with <see cref="PdfCanvas.PaintShading"/> over the current clip: an axial (linear) shading
/// between two points or a radial one between two circles. Coordinates are in the user space in effect when it is painted.
/// </summary>
public sealed class PdfShading
{
    private readonly double[] _coords;
    private readonly IReadOnlyList<ShadingStop> _stops;

    private PdfShading(int type, double[] coords, IReadOnlyList<ShadingStop> stops)
    {
        if (stops.Count < 2) throw new ArgumentException("A shading needs at least two stops.", nameof(stops));
        Type = type;
        _coords = coords;
        _stops = stops;
    }

    internal int Type { get; }

    internal string ResourceName { get; set; } = "";

    /// <summary>A linear shading from (x0, y0) to (x1, y1); colors beyond the ends repeat the end colors.</summary>
    public static PdfShading Axial(double x0, double y0, double x1, double y1, IReadOnlyList<ShadingStop> stops) =>
        new(2, [x0, y0, x1, y1], stops);

    /// <summary>A radial shading from a circle of radius r0 around (x0, y0) to one of radius r1 around (x1, y1).</summary>
    public static PdfShading Radial(double x0, double y0, double r0, double x1, double y1, double r1, IReadOnlyList<ShadingStop> stops) =>
        new(3, [x0, y0, r0, x1, y1, r1], stops);

    internal void Write(PdfWriter writer, int id)
    {
        string N(double v) => PdfFormat.Number(v);
        // One exponential (type 2) function per pair of stops, joined by a stitching (type 3) function.
        var stops = _stops.OrderBy(s => Math.Clamp(s.Position, 0, 1)).Select(s => s with { Position = Math.Clamp(s.Position, 0, 1) }).ToList();
        // The function must cover its whole domain: extend the first and last color to 0 and 1.
        if (stops[0].Position > 0) stops.Insert(0, stops[0] with { Position = 0 });
        if (stops[^1].Position < 1) stops.Add(stops[^1] with { Position = 1 });
        var functionIds = new List<int>();
        for (int i = 0; i + 1 < stops.Count; i++)
        {
            int fid = writer.Allocate();
            writer.WriteObject(fid, $"<< /FunctionType 2 /Domain [0 1] /C0 [{N(stops[i].R)} {N(stops[i].G)} {N(stops[i].B)}] " +
                                    $"/C1 [{N(stops[i + 1].R)} {N(stops[i + 1].G)} {N(stops[i + 1].B)}] /N 1 >>");
            functionIds.Add(fid);
        }
        int stitchId = writer.Allocate();
        var bounds = stops.Skip(1).Take(stops.Count - 2).Select(s => N(Math.Clamp(s.Position, 0, 1)));
        var encode = string.Join(" ", functionIds.Select(_ => "0 1"));
        writer.WriteObject(stitchId,
            $"<< /FunctionType 3 /Domain [0 1] /Functions [{string.Join(" ", functionIds.Select(f => $"{f} 0 R"))}] " +
            $"/Bounds [{string.Join(" ", bounds)}] /Encode [{encode}] >>");
        writer.WriteObject(id,
            $"<< /ShadingType {Type} /ColorSpace /DeviceRGB /Coords [{string.Join(" ", _coords.Select(N))}] /Function {stitchId} 0 R /Extend [true true] >>");
    }
}
