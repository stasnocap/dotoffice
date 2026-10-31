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
}

/// <summary>What a floating object's position is measured from.</summary>
public enum FloatReference
{
    Page, Margin, Column, Character, LeftMargin, RightMargin, InsideMargin, OutsideMargin,
    Paragraph, Line, TopMargin, BottomMargin,
}

/// <summary>Alignment within the reference area; <see cref="Offset"/> uses <see cref="FloatPosition.Offset"/>.</summary>
public enum FloatAlignment { Offset, Start, Center, End }

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

/// <summary>An image anchored to its paragraph and positioned independently of the text flow; size in points.</summary>
public sealed class FloatingImage(ImageData image, double width, double height, RunProperties properties) : Inline(properties)
{
    public ImageData Image { get; } = image;
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
