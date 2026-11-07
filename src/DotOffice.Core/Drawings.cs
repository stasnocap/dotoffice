namespace DotOffice.Core;

/// <summary>A fraction of the frame of the anchor an object belongs to; grouped objects are placed this way inside their group.</summary>
public readonly record struct FrameFraction(double X, double Y, double Width, double Height)
{
    public static FrameFraction Whole { get; } = new(0, 0, 1, 1);

    /// <summary>The fraction <paramref name="inner"/> of this fraction.</summary>
    public FrameFraction Within(FrameFraction inner) => new(X + inner.X * Width, Y + inner.Y * Height, inner.Width * Width, inner.Height * Height);
}

/// <summary>A group an object belongs to that is turned: the group's frame (a fraction of the anchor's) and its rotation in degrees counter-clockwise.</summary>
public sealed record GroupTurn(FrameFraction Frame, double Rotation);

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

    /// <summary>
    /// Rotation in degrees counter-clockwise, including the turns of the groups the object belongs to (which also swing it about their centers,
    /// see <see cref="GroupTurns"/>).
    /// </summary>
    public double Rotation { get; init; }

    /// <summary>The rotated groups the object lies in, outermost first; empty for an object outside groups or in groups that are not turned.</summary>
    public IReadOnlyList<GroupTurn> GroupTurns { get; init; } = [];

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
    /// <summary>A regular pentagon, an octagon, a plus sign, a chevron, a pentagon arrow (<c>homePlate</c>) and a five-pointed star.</summary>
    Pentagon, Octagon, Plus, Chevron, HomePlate, Star5,
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
public sealed record ShapePath(IReadOnlyList<(double X, double Y)> Points, bool Closed, bool Fill = true, bool Stroke = true)
{
    /// <summary>
    /// Further closed rings of the same path (its other subpaths), in the same fractions. They are filled together with <see cref="Points"/> by
    /// the even-odd rule, so a ring inside another one is a hole (a donut), and each is outlined like the first.
    /// </summary>
    public IReadOnlyList<IReadOnlyList<(double X, double Y)>>? Rings { get; init; }
}

public enum ShapeTextAnchor { Top, Middle, Bottom }

/// <summary>A paragraph of the text inside a shape.</summary>
/// <param name="Alignment">How the lines are aligned.</param>
/// <param name="Runs">The text with its formatting.</param>
/// <param name="FontSize">The size of an empty paragraph, or of the largest run.</param>
/// <param name="Indent">Where the paragraph's first line starts, in points from the left inset (DrawingML <c>marL</c> plus the first-line <c>indent</c>).</param>
/// <param name="Margin">Where the following lines start (<c>marL</c>); null starts them at <paramref name="Indent"/>.</param>
/// <param name="LineSpacing">The line pitch as a multiple of single spacing (<c>lnSpc/spcPct</c>).</param>
/// <param name="LinePitch">An exact line pitch in points (<c>lnSpc/spcPts</c>), used instead of <paramref name="LineSpacing"/>.</param>
/// <param name="SpaceBefore">Space above the paragraph in points (<c>spcBef</c>).</param>
/// <param name="SpaceAfter">Space below the paragraph in points (<c>spcAft</c>).</param>
public sealed record ShapeParagraph(HorizontalAlignment Alignment, IReadOnlyList<CellTextRun> Runs, double FontSize, double Indent = 0, double? Margin = null,
    double LineSpacing = 1, double? LinePitch = null, double SpaceBefore = 0, double SpaceAfter = 0);

public enum HatchKind { Horizontal, Vertical, UpDiagonal, DownDiagonal, Cross, DiagonalCross, Tint }

/// <summary>
/// A pattern fill: a <paramref name="Background"/> hatched with lines of the <paramref name="Foreground"/> color every
/// <paramref name="Spacing"/> points, <paramref name="Weight"/> points thick. <see cref="HatchKind.Tint"/> patterns (dots and the like)
/// are drawn as the foreground blended into the background by <paramref name="Coverage"/>.
/// </summary>
public sealed record ShapePattern(HatchKind Kind, Color Foreground, Color Background, double Spacing = 6, double Weight = 1, double Coverage = 0.5);

/// <summary>
/// A shadow drawn as a copy of the shape moved by (<paramref name="Dx"/>, <paramref name="Dy"/>) points, painted with <paramref name="Opacity"/>;
/// with a <paramref name="Blur"/> radius (points) it is several translucent copies spread around that place.
/// </summary>
public sealed record ShapeShadow(Color Color, double Opacity, double Dx, double Dy, double Blur = 0);

/// <summary>
/// A picture fill repeated over the shape: tiles of <paramref name="Width"/> x <paramref name="Height"/> points, shifted by the offsets, starting
/// from the corner or edge <paramref name="Alignment"/> names (tl, t, tr, l, ctr, r, bl, b, br).
/// </summary>
public sealed record ShapeTile(double Width, double Height, double OffsetX, double OffsetY, string Alignment = "tl", string Flip = "none");

/// <summary>A glow: the shape's outline spread by <paramref name="Radius"/> points in a translucent color.</summary>
public sealed record ShapeGlow(Color Color, double Opacity, double Radius);

/// <summary>
/// A reflection: the shape mirrored below its bottom edge (<paramref name="Distance"/> points further down), fading from
/// <paramref name="StartOpacity"/> at the edge to <paramref name="EndOpacity"/> at <paramref name="EndPosition"/> (a fraction of the shape's height).
/// </summary>
public sealed record ShapeReflection(double StartOpacity, double EndOpacity, double EndPosition, double Distance);

/// <summary>A shape, text box or line of a drawing.</summary>
public sealed record SheetShape : SheetObject
{
    public ShapeGeometry Geometry { get; init; }

    /// <summary>Corner radius of rounded rectangles as a fraction of the shorter side (0.1667 by default in Excel).</summary>
    public double CornerRadius { get; init; } = 1.0 / 6;

    /// <summary>
    /// The shape's first and second adjust values as fractions (the preset's <c>adj</c> or <c>adj1</c> over 100000, and <c>adj2</c>); null
    /// keeps the preset's default. What they mean depends on the geometry: the slant of a parallelogram or trapezoid, the depth of a chevron's
    /// or pentagon arrow's point, the cut corner of an octagon, the arm width of a plus, and the shaft width and head length of a block arrow.
    /// </summary>
    public double? Adjust1 { get; init; }

    public double? Adjust2 { get; init; }

    /// <summary>The fill; null means none.</summary>
    public Color? Fill { get; init; }

    /// <summary>A gradient fill; <see cref="Fill"/> then holds the color of its first stop.</summary>
    public GradientFill? FillGradient { get; init; }

    /// <summary>A pattern fill (hatching); drawn instead of <see cref="Fill"/>, which holds its foreground color.</summary>
    public ShapePattern? FillPattern { get; init; }

    /// <summary>A picture fill, stretched over the shape's frame and cut at its outline; drawn instead of <see cref="Fill"/>.</summary>
    public ImageData? FillImage { get; init; }

    /// <summary>When set, <see cref="FillImage"/> is repeated at this size instead of being stretched.</summary>
    public ShapeTile? FillTile { get; init; }

    /// <summary>How opaque the picture fill is (the blip's alphaModFix); 1 is fully opaque.</summary>
    public double FillImageAlpha { get; init; } = 1;

    /// <summary>A glow around the outline; null means none.</summary>
    public ShapeGlow? Glow { get; init; }

    public ShapeReflection? Reflection { get; init; }

    /// <summary>Radius in points over which the edges of the shape fade out (a soft-edge effect); 0 means hard edges.</summary>
    public double SoftEdge { get; init; }

    /// <summary>An outer shadow; null means none.</summary>
    public ShapeShadow? Shadow { get; init; }

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

    /// <summary>
    /// Whether the shape grows taller to fit its text (DrawingML <c>spAutoFit</c>); its top stays where it is. Applied when slides are laid out.
    /// </summary>
    public bool GrowToFit { get; init; }

    /// <summary>The rows of a table this shape is a cell or border of; null for a shape outside tables. Rows grow to fit their cells' text.</summary>
    public TableRowSpan? TableRows { get; init; }
}

/// <summary>
/// Where a cell or a border line lies in a table drawn as separate shapes: from row <paramref name="First"/> over <paramref name="Count"/> rows
/// (a horizontal border has no rows and sits at the top of <paramref name="First"/>). <paramref name="RowHeights"/> are the heights of all the
/// table's rows in any one unit, and <paramref name="Table"/> tells the tables of a slide apart.
/// </summary>
public sealed record TableRowSpan(int Table, int First, int Count, IReadOnlyList<double> RowHeights);
