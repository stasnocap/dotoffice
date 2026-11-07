using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Layout;

/// <summary>
/// Result of layout: pages of positioned primitives. Coordinates are points with the origin at the
/// <b>top-left</b> corner of the page and y growing downwards; renderers convert once.
/// </summary>
public sealed class LayoutDocument
{
    public List<LayoutPage> Pages { get; } = [];
    public string? Title { get; init; }
    public string? Author { get; init; }
}

public sealed class LayoutPage(double width, double height)
{
    public double Width { get; } = width;
    public double Height { get; } = height;

    /// <summary>Items in paint order (backgrounds first).</summary>
    public List<LayoutItem> Items { get; } = [];
}

public abstract record LayoutItem
{
    /// <summary>Returns a copy moved by (dx, dy).</summary>
    public abstract LayoutItem Offset(double dx, double dy);
}

/// <summary>A run of glyphs from one font on one baseline, starting at <see cref="X"/>.</summary>
public sealed record TextItem(double X, double Baseline, TrueTypeFont Font, double FontSize, string Text, Color Color,
    bool SyntheticBold = false, bool SyntheticItalic = false, double CharacterSpacing = 0) : LayoutItem
{
    /// <summary>Set for dynamic fields whose text is substituted after pagination.</summary>
    public FieldRun? Field { get; init; }

    /// <summary>
    /// Shaped glyphs in logical order (positions in font units). When set they are drawn instead of mapping
    /// <see cref="Text"/> through the cmap; right-to-left glyphs are drawn in reverse order.
    /// </summary>
    public IReadOnlyList<ShapedGlyph>? Glyphs { get; init; }

    public bool RightToLeft { get; init; }

    public override LayoutItem Offset(double dx, double dy) => this with { X = X + dx, Baseline = Baseline + dy };
}

public enum LineStyle { Solid, Dotted, Dashed }

public sealed record LineItem(double X1, double Y1, double X2, double Y2, double Width, Color Color, LineStyle Style = LineStyle.Solid)
    : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { X1 = X1 + dx, X2 = X2 + dx, Y1 = Y1 + dy, Y2 = Y2 + dy };
}

public sealed record RectItem(double X, double Y, double Width, double Height, Color Fill) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { X = X + dx, Y = Y + dy };
}

/// <summary>A filled closed polygon (used for icons).</summary>
public sealed record PolygonItem(IReadOnlyList<(double X, double Y)> Points, Color Fill) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { Points = Points.Select(p => (p.X + dx, p.Y + dy)).ToList() };
}

/// <summary>A filled ellipse inscribed in the given box.</summary>
public sealed record EllipseItem(double X, double Y, double Width, double Height, Color Fill) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { X = X + dx, Y = Y + dy };
}

/// <summary>A rectangle filled with a linear gradient; <see cref="Degree"/> runs clockwise from left-to-right (90 = top to bottom).</summary>
public sealed record GradientItem(double X, double Y, double Width, double Height, double Degree, IReadOnlyList<GradientStop> Stops) : LayoutItem
{
    /// <summary>When set, the gradient is painted only inside this polygon (which lies within the box) instead of the whole box.</summary>
    public IReadOnlyList<(double X, double Y)>? Clip { get; init; }

    /// <summary>Whether the colors blend outwards from the focus (<see cref="FocusX"/>, <see cref="FocusY"/>, fractions of the box) instead of along <see cref="Degree"/>.</summary>
    public bool Radial { get; init; }

    public double FocusX { get; init; } = 0.5;

    public double FocusY { get; init; } = 0.5;

    /// <summary>With <see cref="Radial"/>: the colors run in concentric rectangles from the focus to the box's edges instead of in circles.</summary>
    public bool Rectangular { get; init; }

    public override LayoutItem Offset(double dx, double dy) =>
        this with { X = X + dx, Y = Y + dy, Clip = Clip?.Select(p => (p.X + dx, p.Y + dy)).ToList() };
}

public enum PathPointKind { Move, Line, Curve, Close }

/// <summary>A point of a <see cref="PathItem"/>: three <see cref="PathPointKind.Curve"/> points in a row make a cubic Bézier segment; a close point has no position.</summary>
public readonly record struct PathPoint(double X, double Y, PathPointKind Kind);

/// <summary>An outline of straight and Bézier segments, filled and/or stroked (used for drawings in metafiles).</summary>
public sealed record PathItem(IReadOnlyList<PathPoint> Points, Color? Fill, Color? Stroke, double LineWidth = 1, LineStyle Style = LineStyle.Solid,
    bool EvenOdd = false) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { Points = Points.Select(p => p with { X = p.X + dx, Y = p.Y + dy }).ToList() };
}

public sealed record ImageItem(double X, double Y, double Width, double Height, ImageData Image) : LayoutItem
{
    /// <summary>When set, the picture is painted only inside this polygon (a picture fill of a shape).</summary>
    public IReadOnlyList<(double X, double Y)>? Clip { get; init; }

    public override LayoutItem Offset(double dx, double dy) =>
        this with { X = X + dx, Y = Y + dy, Clip = Clip?.Select(p => (p.X + dx, p.Y + dy)).ToList() };
}

/// <summary>
/// Starts a group: everything up to the matching <see cref="PopItem"/> is clipped to <see cref="Clip"/> (when set) and rotated by
/// <see cref="Rotation"/> degrees counter-clockwise about (<see cref="CenterX"/>, <see cref="CenterY"/>), and mirrored about that center when
/// <see cref="FlipH"/> / <see cref="FlipV"/> are set (the mirroring is applied first). Everything in the group is painted with <see cref="Opacity"/>
/// (1 is opaque). Groups nest.
/// </summary>
public sealed record PushItem((double X, double Y, double Width, double Height)? Clip, double Rotation = 0, double CenterX = 0, double CenterY = 0, bool FlipH = false, bool FlipV = false, double Opacity = 1)
    : LayoutItem
{
    /// <summary>
    /// A further clip by outlines (points that are NaN separate the rings): the group is painted where an odd number of the rings cover it, so
    /// a big ring around a shape's outline clips to the outside of the shape.
    /// </summary>
    public IReadOnlyList<(double X, double Y)>? ClipOutline { get; init; }

    public override LayoutItem Offset(double dx, double dy) =>
        this with { Clip = Clip is { } c ? (c.X + dx, c.Y + dy, c.Width, c.Height) : null, CenterX = CenterX + dx, CenterY = CenterY + dy,
                    ClipOutline = ClipOutline?.Select(p => (p.X + dx, p.Y + dy)).ToList() };
}

/// <summary>Ends the innermost <see cref="PushItem"/> group.</summary>
public sealed record PopItem : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this;
}

/// <summary>A clickable area linking to a URI or to a bookmark (<see cref="AnchorItem"/>).</summary>
public sealed record LinkItem(double X, double Y, double Width, double Height, Hyperlink Link) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { X = X + dx, Y = Y + dy };
}

/// <summary>The position of a bookmark (top of its line); paints nothing.</summary>
public sealed record AnchorItem(string Name, double X, double Y) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { X = X + dx, Y = Y + dy };
}

public sealed class LayoutOptions
{
    public required FontResolver Fonts { get; init; }
    public IConversionWarningSink? Warnings { get; init; }
}