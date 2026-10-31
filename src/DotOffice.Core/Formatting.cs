namespace DotOffice.Core;

public enum UnderlineStyle { None, Single, Double, Thick, Dotted, Dashed, Wavy, Words }

public enum VerticalPosition { Baseline, Superscript, Subscript }

/// <summary>Fully resolved character formatting. All lengths in points.</summary>
public sealed record RunProperties
{
    public static RunProperties Default { get; } = new();

    /// <summary>Font for ASCII characters.</summary>
    public string FontFamily { get; init; } = "Times New Roman";

    /// <summary>Font for non-ASCII Latin/Cyrillic/Greek text; null means <see cref="FontFamily"/>.</summary>
    public string? HighAnsiFontFamily { get; init; }

    /// <summary>Font for CJK text; null means <see cref="HighAnsiFontFamily"/>.</summary>
    public string? EastAsianFontFamily { get; init; }

    public double FontSize { get; init; } = 10;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public UnderlineStyle Underline { get; init; }
    public bool Strike { get; init; }
    public bool DoubleStrike { get; init; }

    /// <summary>Text color; null means automatic (black on light backgrounds).</summary>
    public Color? Color { get; init; }

    public Color? Highlight { get; init; }

    /// <summary>Run background shading (drawn like highlight, below it).</summary>
    public Color? Shading { get; init; }

    public VerticalPosition VerticalPosition { get; init; }
    public bool Caps { get; init; }
    public bool SmallCaps { get; init; }
    public bool Hidden { get; init; }

    /// <summary>Extra space added after every character.</summary>
    public double CharacterSpacing { get; init; }

    /// <summary>Baseline raise (positive) or lowering (negative), independent of super/subscript.</summary>
    public double BaselineShift { get; init; }

    /// <summary>Font for complex-script text (Arabic, Hebrew…); null means <see cref="FontFamily"/>.</summary>
    public string? ComplexScriptFontFamily { get; init; }

    /// <summary>Size, bold and italic of complex-script text; null means the same as other text.</summary>
    public double? ComplexScriptFontSize { get; init; }
    public bool? ComplexScriptBold { get; init; }
    public bool? ComplexScriptItalic { get; init; }

    /// <summary>The run is right-to-left: its neutral characters and digits read right to left.</summary>
    public bool RightToLeft { get; init; }

    /// <summary>Format all characters of the run with the complex-script properties.</summary>
    public bool ComplexScript { get; init; }

    /// <summary>Apply font kerning at this size and above (null: no kerning, as in Word by default).</summary>
    public double? KerningMinSize { get; init; }
}

public enum ParagraphAlignment { Left, Center, Right, Justify, Distribute }

public enum LineSpacingRule
{
    /// <summary><see cref="LineSpacing.Value"/> is a multiple of the font's natural line height.</summary>
    Auto,
    /// <summary><see cref="LineSpacing.Value"/> is the exact line height in points.</summary>
    Exact,
    /// <summary><see cref="LineSpacing.Value"/> is the minimum line height in points.</summary>
    AtLeast,
}

public readonly record struct LineSpacing(double Value, LineSpacingRule Rule)
{
    public static LineSpacing Single { get; } = new(1, LineSpacingRule.Auto);
}

public enum TabAlignment { Left, Center, Right, Decimal, Bar }

public enum TabLeader { None, Dot, Hyphen, Underscore, MiddleDot }

/// <summary>A tab stop; <see cref="Position"/> is measured from the start of the text area (left margin).</summary>
public readonly record struct TabStop(double Position, TabAlignment Alignment, TabLeader Leader = TabLeader.None);

public enum BorderStyle { None, Single, Double, Dotted, Dashed, Thick }

/// <summary>One border line. <see cref="Space"/> is the gap between border and content.</summary>
public readonly record struct Border(BorderStyle Style, double Width, Color? Color, double Space = 0)
{
    public static Border None { get; } = new(BorderStyle.None, 0, null);
    public bool IsVisible => Style != BorderStyle.None && Width > 0;
}

/// <summary>Borders for a paragraph, table or cell. <see cref="InsideHorizontal"/> doubles as the
/// "between" border for paragraph groups.</summary>
public sealed record BorderSet
{
    public static BorderSet Empty { get; } = new();

    public Border Top { get; init; } = Border.None;
    public Border Bottom { get; init; } = Border.None;
    public Border Left { get; init; } = Border.None;
    public Border Right { get; init; } = Border.None;
    public Border InsideHorizontal { get; init; } = Border.None;
    public Border InsideVertical { get; init; } = Border.None;

    public bool HasAny => Top.IsVisible || Bottom.IsVisible || Left.IsVisible || Right.IsVisible
                          || InsideHorizontal.IsVisible || InsideVertical.IsVisible;
}

/// <summary>Fully resolved paragraph formatting. All lengths in points.</summary>
public sealed record ParagraphProperties
{
    public static ParagraphProperties Default { get; } = new();

    /// <summary>Opaque style identifier, used for contextual spacing between same-style paragraphs.</summary>
    public string? StyleId { get; init; }

    public ParagraphAlignment Alignment { get; init; }
    public double LeftIndent { get; init; }
    public double RightIndent { get; init; }

    /// <summary>First-line offset relative to <see cref="LeftIndent"/>; negative values are hanging indents.</summary>
    public double FirstLineIndent { get; init; }

    public double SpaceBefore { get; init; }
    public double SpaceAfter { get; init; }
    public LineSpacing LineSpacing { get; init; } = LineSpacing.Single;

    /// <summary>Ignore spacing between consecutive paragraphs with the same <see cref="StyleId"/>.</summary>
    public bool ContextualSpacing { get; init; }

    public bool KeepWithNext { get; init; }
    public bool KeepLinesTogether { get; init; }
    public bool PageBreakBefore { get; init; }
    public bool WidowControl { get; init; } = true;

    /// <summary>Custom tab stops sorted by position.</summary>
    public IReadOnlyList<TabStop> TabStops { get; init; } = [];

    public BorderSet Borders { get; init; } = BorderSet.Empty;
    public Color? Shading { get; init; }

    /// <summary>
    /// Right-to-left paragraph: lines start at the right, and alignment, indents and tab stops are measured
    /// from the right edge (<see cref="ParagraphAlignment.Left"/> and <see cref="LeftIndent"/> mean the start side).
    /// </summary>
    public bool RightToLeft { get; init; }
}