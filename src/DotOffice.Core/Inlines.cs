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