namespace DotOffice.Core;

public abstract class Inline(RunProperties properties)
{
    public RunProperties Properties { get; } = properties;
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

public enum FieldKind { Page, NumPages }

/// <summary>A dynamic field evaluated during layout.</summary>
public sealed class FieldRun(FieldKind kind, RunProperties properties) : Inline(properties)
{
    public FieldKind Kind { get; } = kind;
}