namespace DotOffice.Core;

/// <summary>A fraction of the frame of the anchor an object belongs to; grouped objects are placed this way inside their group.</summary>
public readonly record struct FrameFraction(double X, double Y, double Width, double Height)
{
    public static FrameFraction Whole { get; } = new(0, 0, 1, 1);

    /// <summary>The fraction <paramref name="inner"/> of this fraction.</summary>
    public FrameFraction Within(FrameFraction inner) => new(X + inner.X * Width, Y + inner.Y * Height, inner.Width * Width, inner.Height * Height);
}

/// <summary>An object placed on a worksheet by a drawing anchor (a picture or a shape). Lengths in points.</summary>
public abstract record SheetObject
{
    public ImageAnchorKind Anchor { get; init; }

    public AnchorPoint From { get; init; }

    public AnchorPoint To { get; init; }

    public double X { get; init; }

    public double Y { get; init; }

    public double Width { get; init; }

    public double Height { get; init; }

    /// <summary>Where the object lies inside the frame the anchor describes; the whole frame unless the object is part of a group.</summary>
    public FrameFraction Frame { get; init; } = FrameFraction.Whole;

    /// <summary>Rotation in degrees counter-clockwise.</summary>
    public double Rotation { get; init; }

    /// <summary>Mirrored left to right / top to bottom (about the center of the frame, before the rotation is applied).</summary>
    public bool FlipH { get; init; }
    public bool FlipV { get; init; }

    /// <summary>Position in the drawing's z-order (later is on top), shared by all kinds of objects of a sheet.</summary>
    public int ZOrder { get; init; }
}

/// <summary>A picture placed on a worksheet by a drawing anchor.</summary>
public sealed record SheetImage : SheetObject
{
    public required ImageData Image { get; init; }

    /// <summary>Cropping as fractions of the image removed from each side.</summary>
    public double CropLeft { get; init; }
    public double CropTop { get; init; }
    public double CropRight { get; init; }
    public double CropBottom { get; init; }
}

public enum ShapeGeometry
{
    Rectangle, RoundRectangle, Ellipse, Triangle, RightTriangle, Diamond, Parallelogram, Trapezoid, Hexagon,
    RightArrow, LeftArrow, UpArrow, DownArrow,
    /// <summary>A straight line from the top-left to the bottom-right of the frame (flips choose the other diagonal).</summary>
    Line,
    /// <summary>Custom geometry: see <see cref="SheetShape.Paths"/>.</summary>
    Custom,
}

public enum StrokeStyle { Solid, Dotted, Dashed }

public enum LineEndKind { None, Triangle, Stealth, Arrow, Oval, Diamond }

/// <summary>The end of a line: the kind of arrowhead and its width and length as multiples of the line's width.</summary>
public readonly record struct LineEnd(LineEndKind Kind, double Width = 3, double Length = 3);

/// <summary>
/// One path of a shape drawn as lines through its <paramref name="Points"/>, which are fractions of the shape's frame (curves are already
/// sampled). Closed paths are filled; <paramref name="Fill"/> and <paramref name="Stroke"/> switch the fill and the outline off.
/// </summary>
public sealed record ShapePath(IReadOnlyList<(double X, double Y)> Points, bool Closed, bool Fill = true, bool Stroke = true);

public enum ShapeTextAnchor { Top, Middle, Bottom }

/// <summary>A paragraph of the text inside a shape.</summary>
public sealed record ShapeParagraph(HorizontalAlignment Alignment, IReadOnlyList<CellTextRun> Runs, double FontSize);

/// <summary>A shape, text box or line of a drawing.</summary>
public sealed record SheetShape : SheetObject
{
    public ShapeGeometry Geometry { get; init; }

    /// <summary>Corner radius of rounded rectangles as a fraction of the shorter side (0.1667 by default in Excel).</summary>
    public double CornerRadius { get; init; } = 1.0 / 6;

    /// <summary>The fill; null means none.</summary>
    public Color? Fill { get; init; }

    /// <summary>A gradient fill; <see cref="Fill"/> then holds the color of its first stop.</summary>
    public GradientFill? FillGradient { get; init; }

    /// <summary>
    /// Paths that replace the preset geometry: custom geometry and bent or curved connectors. Open paths of a line are the connector's route.
    /// </summary>
    public IReadOnlyList<ShapePath>? Paths { get; init; }

    /// <summary>The arrowheads at the start and at the end of a line or open path.</summary>
    public LineEnd HeadEnd { get; init; }

    public LineEnd TailEnd { get; init; }

    /// <summary>The outline color; null means no outline.</summary>
    public Color? LineColor { get; init; }

    public double LineWidth { get; init; } = 1;

    public StrokeStyle LineStyle { get; init; }

    public List<ShapeParagraph> Paragraphs { get; init; } = [];

    public ShapeTextAnchor TextAnchor { get; init; }

    /// <summary>Space between the frame and the text.</summary>
    public Thickness TextInsets { get; init; } = new(3.6, 7.2, 3.6, 7.2);

    /// <summary>Whether text wraps at the right edge of the text area.</summary>
    public bool WrapText { get; init; } = true;
}
