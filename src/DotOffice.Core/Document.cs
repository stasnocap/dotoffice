namespace DotOffice.Core;

/// <summary>A format-neutral flow document: a sequence of sections holding blocks.</summary>
public sealed class Document
{
    public List<Section> Sections { get; } = [];
    public DocumentSettings Settings { get; init; } = new();
    public string? Title { get; set; }
    public string? Author { get; set; }
}

public sealed class DocumentSettings
{
    /// <summary>Distance between implicit tab stops.</summary>
    public double DefaultTabStop { get; init; } = 36;

    /// <summary>Use the even header/footer on even-numbered pages.</summary>
    public bool EvenAndOddHeaders { get; init; }

    /// <summary>
    /// When true, a table's left edge is placed at its indent; when false (legacy Word behavior) the table
    /// is shifted left by its first cell's left margin so cell text lines up with surrounding paragraphs.
    /// </summary>
    public bool TableEdgeAtIndent { get; init; } = true;

    /// <summary>When false, lines ending in a manual line break are not stretched in justified paragraphs.</summary>
    public bool JustifyLinesEndingWithBreak { get; init; } = true;
}

public enum Orientation { Portrait, Landscape }

public enum SectionStart { NextPage, Continuous, EvenPage, OddPage }

public enum PageNumberFormat { Decimal, UpperRoman, LowerRoman, UpperLetter, LowerLetter }

public sealed class Section
{
    public double PageWidth { get; init; } = 612;
    public double PageHeight { get; init; } = 792;
    public Orientation Orientation { get; init; }
    public Thickness Margins { get; init; } = new(72, 72, 72, 72);

    /// <summary>Distance from the top page edge to the top of the header.</summary>
    public double HeaderDistance { get; init; } = 36;

    /// <summary>Distance from the bottom page edge to the bottom of the footer.</summary>
    public double FooterDistance { get; init; } = 36;

    public SectionStart Start { get; init; }
    public bool DifferentFirstPage { get; init; }
    public int? PageNumberStart { get; init; }
    public PageNumberFormat PageNumberFormat { get; init; }

    public HeaderFooter? DefaultHeader { get; init; }
    public HeaderFooter? FirstHeader { get; init; }
    public HeaderFooter? EvenHeader { get; init; }
    public HeaderFooter? DefaultFooter { get; init; }
    public HeaderFooter? FirstFooter { get; init; }
    public HeaderFooter? EvenFooter { get; init; }

    public SectionColumns Columns { get; init; } = SectionColumns.Single;

    public List<Block> Blocks { get; } = [];

    public double ContentWidth => PageWidth - Margins.Left - Margins.Right;
}

/// <summary>One text column of explicit width, followed by <see cref="SpaceAfter"/> before the next column.</summary>
public readonly record struct ColumnDefinition(double Width, double SpaceAfter);

/// <summary>Text columns of a section.</summary>
public sealed record SectionColumns
{
    public static SectionColumns Single { get; } = new();

    public int Count { get; init; } = 1;

    /// <summary>Gap between equal-width columns.</summary>
    public double Spacing { get; init; } = 36;

    /// <summary>Explicit columns; when empty, <see cref="Count"/> equal columns separated by <see cref="Spacing"/> are used.</summary>
    public IReadOnlyList<ColumnDefinition> Definitions { get; init; } = [];

    /// <summary>Draw a vertical line between columns.</summary>
    public bool Separator { get; init; }

    /// <summary>Column x offsets (relative to the left margin) and widths for a text area of <paramref name="contentWidth"/>.</summary>
    public IReadOnlyList<(double X, double Width)> Resolve(double contentWidth)
    {
        if (Definitions.Count > 1)
        {
            double total = Definitions.Sum(d => d.Width) + Definitions.Take(Definitions.Count - 1).Sum(d => d.SpaceAfter);
            double scale = total > contentWidth && total > 0 ? contentWidth / total : 1;
            var explicitColumns = new List<(double, double)>(Definitions.Count);
            double x = 0;
            foreach (var d in Definitions)
            {
                explicitColumns.Add((x, d.Width * scale));
                x += (d.Width + d.SpaceAfter) * scale;
            }
            return explicitColumns;
        }
        int count = Math.Clamp(Count, 1, 45);
        double spacing = Math.Clamp(Spacing, 0, contentWidth / count);
        double width = (contentWidth - spacing * (count - 1)) / count;
        if (count == 1 || width < 9) return [(0, contentWidth)];
        return Enumerable.Range(0, count).Select(i => (i * (width + spacing), width)).ToList();
    }
}

/// <summary>Header or footer content (a small block story).</summary>
public sealed class HeaderFooter
{
    public List<Block> Blocks { get; } = [];
}