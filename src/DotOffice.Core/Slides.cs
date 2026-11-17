namespace DotOffice.Core;

/// <summary>A presentation: its slides in show order, hidden slides left out.</summary>
public sealed class Presentation
{
    public List<Slide> Slides { get; } = [];

    public string? Title { get; set; }

    public string? Author { get; set; }
}

/// <summary>
/// A slide of a presentation, one page in points. Its objects use <see cref="ImageAnchorKind.Absolute"/> placement measured from the slide's
/// top-left corner.
/// </summary>
public sealed class Slide(double width, double height)
{
    public double Width { get; } = width;

    public double Height { get; } = height;

    /// <summary>The background color; null leaves the page white.</summary>
    public Color? Background { get; set; }

    /// <summary>A gradient background, drawn instead of <see cref="Background"/>.</summary>
    public GradientFill? BackgroundGradient { get; set; }

    /// <summary>A picture background stretched over the slide, drawn instead of <see cref="Background"/>.</summary>
    public ImageData? BackgroundImage { get; set; }

    /// <summary>The bookmark name that links to this slide (<see cref="Hyperlink.Bookmark"/>) jump to; null when nothing links here.</summary>
    public string? Anchor { get; set; }

    /// <summary>Pictures and shapes, drawn in <see cref="SheetObject.ZOrder"/> order.</summary>
    public List<SheetObject> Objects { get; } = [];
}
