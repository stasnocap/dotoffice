namespace DotOffice.Core;

public abstract class Inline(RunProperties properties)
{
    public RunProperties Properties { get; } = properties;

    /// <summary>Hyperlink target when the inline is part of a link.</summary>
    public Hyperlink? Link { get; set; }
}

/// <summary>A link to an external <see cref="Uri"/> or to a <see cref="Bookmark"/> in the document.</summary>
public sealed record Hyperlink(string? Uri, string? Bookmark = null);

/// <summary>A named position in the text (link destination, PAGEREF/REF target). Takes no space.</summary>
public sealed class BookmarkStart(string name, RunProperties properties) : Inline(properties)
{
    public string Name { get; } = name;
}

public sealed class TextRun(string text, RunProperties properties) : Inline(properties)
{
    public string Text { get; } = text;
}

/// <summary>An image placed inline with text; size in points.</summary>
public sealed class ImageRun(ImageData image, double width, double height, RunProperties properties) : Inline(properties)
{
    public ImageData Image { get; } = image;
    public double Width { get; } = width;
    public double Height { get; } = height;

    /// <summary>Set when the run is a chart (the image is then empty): the chart to draw in the run's box, or null when it could not be read and its space stays empty.</summary>
    public ChartContent? Chart { get; init; }
}

/// <summary>A chart of a document, read from its chart part; <see cref="Chart"/> is null when nothing can be drawn (the box stays empty).</summary>
public sealed record ChartContent(SheetChart? Chart);

/// <summary>What a floating object's position is measured from.</summary>
public enum FloatReference
{
    Page, Margin, Column, Character, LeftMargin, RightMargin, InsideMargin, OutsideMargin,
    Paragraph, Line, TopMargin, BottomMargin,
}

/// <summary>Alignment within the reference area; <see cref="Offset"/> uses <see cref="FloatPosition.Offset"/>.</summary>
public enum FloatAlignment { Offset, Start, Center, End, Inside, Outside }

/// <summary>One axis of a floating object's position; <see cref="Offset"/> is in points from the reference's start.</summary>
public readonly record struct FloatPosition(FloatReference RelativeTo, FloatAlignment Alignment, double Offset);

public enum TextWrap
{
    /// <summary>Text ignores the object (it is drawn in front of or behind the text).</summary>
    None,
    /// <summary>No text beside the object: lines continue below it.</summary>
    TopAndBottom,
    /// <summary>Text flows around the object's bounding rectangle.</summary>
    Square,
}

/// <summary>An object anchored to its paragraph and positioned independently of the text flow; size in points.</summary>
public abstract class FloatingObject(double width, double height, RunProperties properties) : Inline(properties)
{
    public double Width { get; } = width;
    public double Height { get; } = height;
    public FloatPosition Horizontal { get; init; } = new(FloatReference.Column, FloatAlignment.Offset, 0);
    public FloatPosition Vertical { get; init; } = new(FloatReference.Paragraph, FloatAlignment.Offset, 0);
    public TextWrap Wrap { get; init; }

    /// <summary>Drawn below the text (otherwise above it).</summary>
    public bool BehindText { get; init; }

    /// <summary>Minimum distance between the image and wrapped text.</summary>
    public Thickness Distance { get; init; }
}

/// <summary>A picture anchored to its paragraph.</summary>
public sealed class FloatingImage(ImageData image, double width, double height, RunProperties properties) : FloatingObject(width, height, properties)
{
    public ImageData Image { get; } = image;

    /// <summary>Set when the object is a chart (the image is then empty); see <see cref="ImageRun.Chart"/>.</summary>
    public ChartContent? Chart { get; init; }
}

public enum TextBoxAnchor { Top, Center, Bottom }

/// <summary>Direction of the text of a text box (<c>bodyPr vert</c>): turned a quarter clockwise (<c>vert</c>, <c>eaVert</c>) or counter-clockwise (<c>vert270</c>).</summary>
public enum TextBoxDirection { Horizontal, Clockwise, CounterClockwise }

/// <summary>
/// A text box (a DrawingML <c>wps</c> shape with <c>w:txbxContent</c>): a rectangle with optional fill and outline around block content
/// laid out in the box minus <see cref="Insets"/>. Content that does not fit is clipped, as in Word.
/// </summary>
public sealed class FloatingTextBox(double width, double height, RunProperties properties) : FloatingObject(width, height, properties)
{
    public List<Block> Blocks { get; } = [];
    public Color? Fill { get; init; }
    public Color? Outline { get; init; }
    public double OutlineWidth { get; init; } = 0.75;
    public LineStyleKind OutlineStyle { get; init; }
    public Thickness Insets { get; init; } = new(3.6, 7.2, 3.6, 7.2);
    public TextBoxAnchor VerticalAnchor { get; init; }
    public TextBoxDirection TextDirection { get; init; }

    /// <summary>Clockwise rotation in degrees.</summary>
    public double Rotation { get; init; }

    /// <summary>The outline of the box; text goes in the insets of the frame, or of the rectangle inscribed in an ellipse or rounded corners.</summary>
    public ShapeGeometry Geometry { get; init; }

    /// <summary>Corner radius of a rounded rectangle as a fraction of the shorter side.</summary>
    public double CornerRadius { get; init; } = 1.0 / 6;

    /// <summary>The preset's first and second adjust values as fractions; null keeps its default.</summary>
    public double? Adjust1 { get; init; }

    public double? Adjust2 { get; init; }

    /// <summary>A gradient fill, drawn instead of <see cref="Fill"/> (which holds the color of its first stop).</summary>
    public GradientFill? FillGradient { get; init; }
}

/// <summary>A shape laid out in the line of text like a picture (<c>wp:inline</c>); its position fields are not used.</summary>
public sealed class InlineShape(FloatingTextBox box) : Inline(box.Properties)
{
    public FloatingTextBox Box { get; } = box;
    public double Width => Box.Width;
    public double Height => Box.Height;
}

public enum LineStyleKind { Solid, Dotted, Dashed }

public enum BreakKind { Line, Page, Column }

public sealed class Break(BreakKind kind, RunProperties properties) : Inline(properties)
{
    public BreakKind Kind { get; } = kind;
}

public sealed class Tab(RunProperties properties) : Inline(properties);

public enum FieldKind
{
    Page,
    NumPages,

    /// <summary>Number of pages in the current section.</summary>
    SectionPages,

    /// <summary>Page number of <see cref="FieldRun.Bookmark"/>.</summary>
    PageRef,
}

/// <summary>A dynamic field evaluated during layout.</summary>
public sealed class FieldRun(FieldKind kind, RunProperties properties) : Inline(properties)
{
    public FieldKind Kind { get; } = kind;

    /// <summary>Target bookmark of a <see cref="FieldKind.PageRef"/> field.</summary>
    public string? Bookmark { get; init; }

    /// <summary>Number format from a <c>\*</c> switch; null uses the page's section format.</summary>
    public PageNumberFormat? NumberFormat { get; init; }

    /// <summary>Result stored in the document, shown when the field cannot be evaluated.</summary>
    public string? CachedResult { get; set; }
}
public enum NoteKind { Footnote, Endnote }

/// <summary>Footnote or endnote content; its first paragraph normally starts with a <see cref="NoteMark"/>.</summary>
public sealed class Note(NoteKind kind)
{
    public NoteKind Kind { get; } = kind;
    public List<Block> Blocks { get; } = [];

    /// <summary>The mark shown at the reference and in the note, e.g. "1" or "iv"; empty when the author typed a custom mark.</summary>
    public string Mark { get; set; } = "";

    /// <summary>
    /// Set when footnote numbers restart on every page: gives the mark of the n-th (1-based) numbered
    /// footnote on a page. Layout then replaces <see cref="Mark"/> once pages are known.
    /// </summary>
    public Func<int, string>? PageMark { get; init; }
}

/// <summary>A footnote/endnote reference mark in the text (formatted by its run properties, usually superscript).</summary>
public sealed class NoteReference(Note note, RunProperties properties) : Inline(properties)
{
    public Note Note { get; } = note;

    public string Mark => Note.Mark;
}

/// <summary>The number of the note being read, at the start of its text (w:footnoteRef/w:endnoteRef).</summary>
public sealed class NoteMark(Note note, RunProperties properties) : Inline(properties)
{
    public Note Note { get; } = note;
}
