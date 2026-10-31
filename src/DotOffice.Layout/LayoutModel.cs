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

public sealed record ImageItem(double X, double Y, double Width, double Height, ImageData Image) : LayoutItem
{
    public override LayoutItem Offset(double dx, double dy) => this with { X = X + dx, Y = Y + dy };
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