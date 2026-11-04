using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>Parsing of A1-style cell references and the range syntax used by defined names.</summary>
internal static class CellReference
{
    public const int MaxRows = 1_048_576;
    public const int MaxColumns = 16_384;

    /// <summary>Parses "B3" (optionally with $ markers) into zero-based row and column.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out int row, out int column)
    {
        row = column = 0;
        int i = 0;
        if (i < text.Length && text[i] == '$') i++;
        int letters = 0;
        long col = 0;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            if (col > MaxColumns) return false;
            i++;
            letters++;
        }
        if (letters == 0) return false;
        if (i < text.Length && text[i] == '$') i++;
        int digitsStart = i;
        long r = 0;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            r = r * 10 + (text[i] - '0');
            if (r > MaxRows) return false;
            i++;
        }
        if (i == digitsStart || i != text.Length || r < 1) return false;
        row = (int)r - 1;
        column = (int)col - 1;
        return true;
    }

    /// <summary>Parses a column such as "$B" into a zero-based index.</summary>
    public static bool TryParseColumn(ReadOnlySpan<char> text, out int column)
    {
        column = 0;
        if (text.Length > 0 && text[0] == '$') text = text[1..];
        if (text.Length == 0) return false;
        long col = 0;
        foreach (var c in text)
        {
            if (!char.IsAsciiLetter(c)) return false;
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
            if (col > MaxColumns) return false;
        }
        column = (int)col - 1;
        return true;
    }

    /// <summary>Parses a row such as "$3" into a zero-based index.</summary>
    public static bool TryParseRow(ReadOnlySpan<char> text, out int row)
    {
        row = 0;
        if (text.Length > 0 && text[0] == '$') text = text[1..];
        if (text.Length == 0 || !int.TryParse(text, out var value) || value < 1 || value > MaxRows) return false;
        row = value - 1;
        return true;
    }

    /// <summary>Parses "A1" or "A1:C5" (no sheet prefix) into a normalized range.</summary>
    public static bool TryParseRange(ReadOnlySpan<char> text, out CellRange range)
    {
        range = default;
        int colon = text.IndexOf(':');
        if (colon < 0)
        {
            if (!TryParse(text, out var r, out var c)) return false;
            range = new CellRange(r, c, r, c);
            return true;
        }
        if (!TryParse(text[..colon], out var r1, out var c1) || !TryParse(text[(colon + 1)..], out var r2, out var c2))
            return false;
        range = new CellRange(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
        return true;
    }

    /// <summary>
    /// Parses the formula of a defined name such as <c>'Sheet 1'!$A$1:$D$20</c>, <c>Sheet1!$1:$2</c> or <c>Sheet1!$A:$B</c>,
    /// possibly several parts separated by commas. Each part is reported as a cell range, a row span or a column span.
    /// </summary>
    public static IEnumerable<NameArea> ParseAreas(string formula)
    {
        foreach (var part in SplitTopLevel(formula))
        {
            var text = part.AsSpan().Trim();
            int bang = LastBangOutsideQuotes(text);
            string? sheet = null;
            if (bang >= 0)
            {
                sheet = Unquote(text[..bang].ToString());
                text = text[(bang + 1)..];
            }
            int colon = text.IndexOf(':');
            if (colon < 0)
            {
                if (TryParse(text, out var r, out var c)) yield return NameArea.Cells(sheet, new CellRange(r, c, r, c));
                continue;
            }
            var a = text[..colon];
            var b = text[(colon + 1)..];
            if (TryParseRow(a, out var r1) && TryParseRow(b, out var r2))
                yield return NameArea.Rows(sheet, Math.Min(r1, r2), Math.Max(r1, r2));
            else if (TryParseColumn(a, out var c1) && TryParseColumn(b, out var c2) && !char.IsAsciiDigit(a[^1]) && !char.IsAsciiDigit(b[^1]))
                yield return NameArea.Columns(sheet, Math.Min(c1, c2), Math.Max(c1, c2));
            else if (TryParseRange(text, out var range))
                yield return NameArea.Cells(sheet, range);
        }
    }

    private static List<string> SplitTopLevel(string formula)
    {
        var parts = new List<string>();
        int start = 0;
        bool quoted = false;
        for (int i = 0; i < formula.Length; i++)
        {
            if (formula[i] == '\'') quoted = !quoted;
            else if (formula[i] == ',' && !quoted)
            {
                parts.Add(formula[start..i]);
                start = i + 1;
            }
        }
        parts.Add(formula[start..]);
        return parts;
    }

    private static int LastBangOutsideQuotes(ReadOnlySpan<char> text)
    {
        bool quoted = false;
        int found = -1;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == '\'') quoted = !quoted;
            else if (text[i] == '!' && !quoted) found = i;
        }
        return found;
    }

    private static string Unquote(string sheet)
    {
        sheet = sheet.Trim();
        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[^1] == '\'')
            sheet = sheet[1..^1].Replace("''", "'");
        return sheet;
    }
}

/// <summary>One area of a defined name: a cell range, whole rows or whole columns.</summary>
internal readonly record struct NameArea(string? Sheet, CellRange? Range, (int From, int To)? RowSpan, (int From, int To)? ColumnSpan)
{
    public static NameArea Cells(string? sheet, CellRange range) => new(sheet, range, null, null);

    public static NameArea Rows(string? sheet, int from, int to) => new(sheet, null, (from, to), null);

    public static NameArea Columns(string? sheet, int from, int to) => new(sheet, null, null, (from, to));
}
