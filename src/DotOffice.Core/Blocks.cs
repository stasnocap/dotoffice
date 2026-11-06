namespace DotOffice.Core;

public abstract class Block;

public sealed class Paragraph : Block
{
    public ParagraphProperties Properties { get; init; } = ParagraphProperties.Default;

    /// <summary>Formatting of the paragraph mark; determines the height of empty paragraphs.</summary>
    public RunProperties MarkProperties { get; init; } = RunProperties.Default;

    /// <summary>Resolved list number/bullet, or null when the paragraph is not a list item.</summary>
    public ListLabel? ListLabel { get; init; }

    public List<Inline> Inlines { get; } = [];
}

public enum ListLabelSuffix { Tab, Space, Nothing }

/// <summary>A list item's already-formatted label text (e.g. "2.1." or "•").</summary>
public sealed record ListLabel(string Text, RunProperties Properties, ListLabelSuffix Suffix = ListLabelSuffix.Tab,
    ParagraphAlignment Alignment = ParagraphAlignment.Left);

public enum TableWidthUnit { Auto, Points, Percent }

/// <summary>A preferred width; <see cref="TableWidthUnit.Percent"/> values range 0–100.</summary>
public readonly record struct TableWidth(double Value, TableWidthUnit Unit)
{
    public static TableWidth Auto { get; } = new(0, TableWidthUnit.Auto);
}

public enum TableAlignment { Left, Center, Right }

public enum TableLayoutMode { Auto, Fixed }

public sealed record TableProperties
{
    public TableWidth Width { get; init; } = TableWidth.Auto;
    public TableAlignment Alignment { get; init; }

    /// <summary>Offset of the table from the left edge of the text area.</summary>
    public double Indent { get; init; }

    /// <summary>
    /// The table is displayed right to left (<c>w:bidiVisual</c>): its first cell is at the right, and <see cref="Indent"/> and a
    /// <see cref="TableAlignment.Left"/> alignment refer to its leading, right-hand edge.
    /// </summary>
    public bool RightToLeft { get; init; }

    public TableLayoutMode Layout { get; init; }

    /// <summary>Set for a floating table (<c>w:tblpPr</c>); null for a table in the text flow.</summary>
    public TablePosition? Position { get; init; }
}

/// <summary>
/// Where a floating table sits (<c>w:tblpPr</c>) and how close text may come; the positions use the same references as
/// floating pictures (a text anchor is the column horizontally and the following paragraph vertically).
/// </summary>
public sealed record TablePosition
{
    public FloatPosition Horizontal { get; init; } = new(FloatReference.Column, FloatAlignment.Offset, 0);
    public FloatPosition Vertical { get; init; } = new(FloatReference.Paragraph, FloatAlignment.Offset, 0);
    public Thickness Distance { get; init; }
}

public sealed class Table : Block
{
    public TableProperties Properties { get; init; } = new();

    /// <summary>Grid column widths in points (may be empty, in which case layout computes them).</summary>
    public List<double> Grid { get; } = [];

    public List<TableRow> Rows { get; } = [];
}

public enum RowHeightRule { Auto, AtLeast, Exact }

public sealed class TableRow
{
    public double Height { get; init; }
    public RowHeightRule HeightRule { get; init; }

    /// <summary>Repeat this row at the top of each page when the table breaks (leading rows only).</summary>
    public bool IsHeader { get; init; }

    public bool CantSplit { get; init; }

    /// <summary>Number of grid columns skipped before the first cell.</summary>
    public int GridBefore { get; init; }

    public List<TableCell> Cells { get; } = [];
}

public enum VerticalMerge { None, Restart, Continue }

public enum CellVerticalAlignment { Top, Center, Bottom }

public sealed class TableCell
{
    public TableWidth Width { get; init; } = TableWidth.Auto;
    public int GridSpan { get; init; } = 1;
    public VerticalMerge VerticalMerge { get; init; }

    /// <summary>Resolved borders of this cell's four edges.</summary>
    public BorderSet Borders { get; init; } = BorderSet.Empty;

    public Color? Shading { get; init; }
    public Thickness Margins { get; init; } = new(0, 5.4, 0, 5.4);
    public CellVerticalAlignment VerticalAlignment { get; init; }
    public List<Block> Blocks { get; } = [];
}