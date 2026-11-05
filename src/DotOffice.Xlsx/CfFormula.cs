using System.Globalization;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>An empty cell as a formula value.</summary>
internal sealed class CfBlank
{
    public static CfBlank Instance { get; } = new();
}

/// <summary>An error value such as <c>#DIV/0!</c>; errors propagate through formulas and never satisfy a rule.</summary>
internal sealed record CfError(string Code)
{
    public static CfError Value { get; } = new("#VALUE!");
    public static CfError DivisionByZero { get; } = new("#DIV/0!");
    public static CfError NotAvailable { get; } = new("#N/A");
}

internal abstract record CfExpr;
internal sealed record CfConst(object Value) : CfExpr;
/// <summary>A reference as written (zero-based); relative parts move with the cell the formula is evaluated for.</summary>
internal sealed record CfRef(int Row, bool RowAbsolute, int Column, bool ColumnAbsolute, string? Sheet = null) : CfExpr;
internal sealed record CfRange(CfRef From, CfRef To) : CfExpr;
internal sealed record CfUnary(char Op, CfExpr Operand) : CfExpr;
internal sealed record CfBinary(string Op, CfExpr Left, CfExpr Right) : CfExpr;
internal sealed record CfCall(string Name, IReadOnlyList<CfExpr> Args) : CfExpr;

/// <summary>
/// Parser for the formulas that conditional-formatting rules use: arithmetic, comparison, text concatenation, cell references
/// and ranges on this or another sheet, defined names that stand for a reference or a constant, and a fixed set of functions. Anything else throws <see cref="NotSupportedException"/>.
/// </summary>
internal sealed class CfParser
{
    public static readonly HashSet<string> Functions = new(StringComparer.Ordinal)
    {
        "AND", "OR", "NOT", "XOR", "IF", "IFERROR", "MOD", "ROW", "COLUMN", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISEVEN", "ISODD",
        "ABS", "INT", "ROUND", "LEN", "LEFT", "RIGHT", "SEARCH", "FIND", "SUM", "MIN", "MAX", "AVERAGE", "COUNT", "COUNTA", "COUNTIF", "TRUE", "FALSE",
        "SUMPRODUCT", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "OFFSET", "INDIRECT",
    };

    private readonly string _text;
    private int _i;
    private int _depth;

    private CfParser(string text) => _text = text;

    /// <summary>
    /// Looks up a defined name (the workbook's, or the sheet's own) and returns its formula; set around the parsing of one sheet's rules
    /// by the reader, which knows the names.
    /// </summary>
    [ThreadStatic] internal static Func<string, string?>? NameResolver;

    /// <summary>Parses a formula with or without the leading <c>=</c>.</summary>
    public static CfExpr Parse(string formula)
    {
        var parser = new CfParser(formula.Trim().TrimStart('='));
        var expr = parser.Comparison();
        parser.SkipSpaces();
        if (parser._i < parser._text.Length) throw new NotSupportedException($"Unexpected '{parser._text[parser._i]}' in formula.");
        return expr;
    }

    private void SkipSpaces()
    {
        while (_i < _text.Length && char.IsWhiteSpace(_text[_i])) _i++;
    }

    private bool Take(string token)
    {
        SkipSpaces();
        if (string.CompareOrdinal(_text, _i, token, 0, token.Length) != 0) return false;
        _i += token.Length;
        return true;
    }

    private CfExpr Comparison()
    {
        var left = Concat();
        while (true)
        {
            string? op = null;
            foreach (var candidate in new[] { "<=", ">=", "<>", "=", "<", ">" })
                if (Take(candidate)) { op = candidate; break; }
            if (op is null) return left;
            left = new CfBinary(op, left, Concat());
        }
    }

    private CfExpr Concat()
    {
        var left = Additive();
        while (Take("&")) left = new CfBinary("&", left, Additive());
        return left;
    }

    private CfExpr Additive()
    {
        var left = Multiplicative();
        while (true)
        {
            if (Take("+")) left = new CfBinary("+", left, Multiplicative());
            else if (Take("-")) left = new CfBinary("-", left, Multiplicative());
            else return left;
        }
    }

    private CfExpr Multiplicative()
    {
        var left = Power();
        while (true)
        {
            if (Take("*")) left = new CfBinary("*", left, Power());
            else if (Take("/")) left = new CfBinary("/", left, Power());
            else return left;
        }
    }

    private CfExpr Power()
    {
        var left = Unary();
        while (Take("^")) left = new CfBinary("^", left, Unary());
        return left;
    }

    // Excel binds a unary minus tighter than ^ (-2^2 is 4).
    private CfExpr Unary()
    {
        if (Take("-")) return new CfUnary('-', Unary());
        if (Take("+")) return Unary();
        var e = Primary();
        while (Take("%")) e = new CfUnary('%', e);
        return e;
    }

    private CfExpr Primary()
    {
        SkipSpaces();
        if (_i >= _text.Length) throw new NotSupportedException("Unexpected end of formula.");
        char c = _text[_i];
        if (c == '(')
        {
            _i++;
            var inner = Comparison();
            if (!Take(")")) throw new NotSupportedException("Missing ')'.");
            return inner;
        }
        if (c == '"') return new CfConst(ReadString());
        if (c == '\'')
        {
            // A quoted sheet name: 'My sheet'!A1
            var sheet = ReadQuotedSheet();
            if (_i >= _text.Length || _text[_i] != '!') throw new NotSupportedException("A quoted sheet name must be followed by '!'.");
            _i++;
            return ReferenceOn(sheet);
        }
        if (char.IsAsciiDigit(c) || c == '$' && _i + 1 < _text.Length && char.IsAsciiDigit(_text[_i + 1]))
        {
            // A whole-row range such as 1:3 or $2:$2 (digits only on both sides of the colon).
            int j = _i + (c == '$' ? 1 : 0), rowStart = j;
            while (j < _text.Length && char.IsAsciiDigit(_text[j])) j++;
            int afterColon = j + 1 < _text.Length && _text[j] == ':' ? j + 1 + (_text[j + 1] == '$' ? 1 : 0) : -1;
            if (afterColon > 0 && afterColon < _text.Length && char.IsAsciiDigit(_text[afterColon]))
            {
                int k = afterColon;
                while (k < _text.Length && char.IsAsciiDigit(_text[k])) k++;
                if (int.TryParse(_text.AsSpan(rowStart, j - rowStart), out var first) && int.TryParse(_text.AsSpan(afterColon, k - afterColon), out var last) && first >= 1 && last >= 1
                    && first <= CellReference.MaxRows && last <= CellReference.MaxRows)
                {
                    _i = k;
                    return new CfRange(new CfRef(first - 1, true, 0, true), new CfRef(last - 1, true, CellReference.MaxColumns - 1, true));
                }
            }
        }
        if (char.IsAsciiDigit(c) || c == '.') return new CfConst(ReadNumber());
        if (char.IsAsciiLetter(c) || c is '$' or '_')
        {
            int start = _i;
            while (_i < _text.Length && (char.IsAsciiLetterOrDigit(_text[_i]) || _text[_i] is '$' or '_' or '.')) _i++;
            var word = _text[start.._i];
            SkipSpaces();
            if (_i < _text.Length && _text[_i] == '(')
            {
                var name = word.ToUpperInvariant();
                if (!Functions.Contains(name)) throw new NotSupportedException($"The function {name} is not supported.");
                _i++;
                var args = new List<CfExpr>();
                if (!Take(")"))
                {
                    do args.Add(Comparison()); while (Take(","));
                    if (!Take(")")) throw new NotSupportedException("Missing ')'.");
                }
                return new CfCall(name, args);
            }
            if (_i < _text.Length && _text[_i] == '!')
            {
                _i++;
                return ReferenceOn(word);
            }
            if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return new CfConst(true);
            if (word.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return new CfConst(false);
            var from = Reference(word, null);
            if (from is null)
            {
                // A whole-column range: A:A or $B:$D.
                if (_i < _text.Length && _text[_i] == ':' && WholeColumn(word) is { } firstColumn)
                {
                    int save = _i++;
                    int s2 = _i;
                    while (_i < _text.Length && (char.IsAsciiLetter(_text[_i]) || _text[_i] == '$')) _i++;
                    if (WholeColumn(_text[s2.._i]) is { } lastColumn)
                        return new CfRange(new CfRef(0, true, firstColumn.Column, firstColumn.Absolute), new CfRef(CellReference.MaxRows - 1, true, lastColumn.Column, lastColumn.Absolute));
                    _i = save;
                }
                return Name(word);
            }
            return Range(from, null);
        }
        throw new NotSupportedException($"Unexpected '{c}' in formula.");
    }

    /// <summary>A reference or range on <paramref name="sheet"/>, just after the '!'.</summary>
    private CfExpr ReferenceOn(string sheet)
    {
        SkipSpaces();
        int start = _i;
        while (_i < _text.Length && (char.IsAsciiLetterOrDigit(_text[_i]) || _text[_i] == '$')) _i++;
        var word = _text[start.._i];
        var from = Reference(word, sheet) ?? throw new NotSupportedException($"'{word}' is not a cell reference.");
        return Range(from, sheet);
    }

    private CfExpr Range(CfRef from, string? sheet)
    {
        if (_i < _text.Length && _text[_i] == ':')
        {
            _i++;
            int s2 = _i;
            while (_i < _text.Length && (char.IsAsciiLetterOrDigit(_text[_i]) || _text[_i] == '$')) _i++;
            var to = Reference(_text[s2.._i], sheet) ?? throw new NotSupportedException("Unsupported range.");
            return new CfRange(from, to);
        }
        return from;
    }

    /// <summary>A defined name: what it stands for (a reference, a range or a constant) is parsed in its place.</summary>
    private CfExpr Name(string word)
    {
        var resolved = NameResolver?.Invoke(word);
        if (resolved is null) throw new NotSupportedException($"'{word}' is not a cell reference or a defined name.");
        if (_depth >= 5) throw new NotSupportedException($"The name '{word}' refers to itself.");
        var inner = new CfParser(resolved.Trim().TrimStart('=')) { _depth = _depth + 1 };
        var expr = inner.Comparison();
        inner.SkipSpaces();
        if (inner._i < inner._text.Length) throw new NotSupportedException($"The name '{word}' is not a plain reference or value.");
        return expr;
    }

    private string ReadQuotedSheet()
    {
        _i++;
        var sb = new System.Text.StringBuilder();
        while (_i < _text.Length)
        {
            if (_text[_i] == '\'')
            {
                if (_i + 1 < _text.Length && _text[_i + 1] == '\'') { sb.Append('\''); _i += 2; continue; }
                _i++;
                return sb.ToString();
            }
            sb.Append(_text[_i++]);
        }
        throw new NotSupportedException("Unterminated sheet name.");
    }

    /// <summary>The column of a word made of letters only (optionally after a '$'), as in <c>$A</c>; null for anything else.</summary>
    private static (int Column, bool Absolute)? WholeColumn(string word)
    {
        bool absolute = word.StartsWith('$');
        var letters = absolute ? word[1..] : word;
        if (letters.Length is 0 or > 3 || !letters.All(char.IsAsciiLetter)) return null;
        return CellReference.TryParse(letters + "1", out _, out var column) ? (column, absolute) : null;
    }

    private static CfRef? Reference(string word, string? sheet)
    {
        if (!CellReference.TryParse(word, out var row, out var column)) return null;
        return new CfRef(row, HasRowDollar(word), column, word[0] == '$', sheet);
    }

    // "$A$1" and "A$1" fix the row (a '$' after the column letters); "$A1" does not.
    private static bool HasRowDollar(string word)
    {
        int i = word[0] == '$' ? 1 : 0;
        while (i < word.Length && char.IsAsciiLetter(word[i])) i++;
        return i < word.Length && word[i] == '$';
    }

    private string ReadString()
    {
        _i++;
        var sb = new System.Text.StringBuilder();
        while (_i < _text.Length)
        {
            if (_text[_i] == '"')
            {
                if (_i + 1 < _text.Length && _text[_i + 1] == '"') { sb.Append('"'); _i += 2; continue; }
                _i++;
                return sb.ToString();
            }
            sb.Append(_text[_i++]);
        }
        throw new NotSupportedException("Unterminated string.");
    }

    private double ReadNumber()
    {
        int start = _i;
        while (_i < _text.Length && (char.IsAsciiDigit(_text[_i]) || _text[_i] == '.')) _i++;
        if (_i < _text.Length && _text[_i] is 'e' or 'E')
        {
            int save = _i++;
            if (_i < _text.Length && _text[_i] is '+' or '-') _i++;
            if (_i < _text.Length && char.IsAsciiDigit(_text[_i])) while (_i < _text.Length && char.IsAsciiDigit(_text[_i])) _i++;
            else _i = save;
        }
        return double.TryParse(_text.AsSpan(start, _i - start), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : throw new NotSupportedException("Invalid number.");
    }
}

/// <summary>What a formula can see: cell values by position and which cell it is being evaluated for.</summary>
/// <summary>The cells of another sheet as formulas see them, and the rectangle worth reading.</summary>
internal sealed record CfSheetView(Func<int, int, object> CellAt, CellRange? Bounds);

internal sealed class CfContext(Func<int, int, object> cellAt, int originRow, int originColumn, int row, int column, CellRange bounds, Func<string, CfSheetView?>? sheets = null)
{
    public Func<int, int, object> CellAt { get; } = cellAt;

    /// <summary>Finds another sheet by name; null when the sheet does not exist.</summary>
    public Func<string, CfSheetView?>? Sheets { get; } = sheets;

    public int Row { get; } = row;

    public int Column { get; } = column;

    /// <summary>The cell the formulas were written for (the top-left of the rule's range).</summary>
    public int OriginRow { get; } = originRow;
    public int OriginColumn { get; } = originColumn;

    /// <summary>Ranges are clamped to this rectangle so that whole-column references stay cheap.</summary>
    public CellRange Bounds { get; } = bounds;

    public (int Row, int Column) Resolve(CfRef r) =>
        (r.RowAbsolute ? r.Row : r.Row + Row - OriginRow, r.ColumnAbsolute ? r.Column : r.Column + Column - OriginColumn);
}

internal static class CfEvaluator
{
    private const int MaxRangeCells = 250_000;

    public static object Evaluate(CfExpr expr, CfContext ctx)
    {
        switch (expr)
        {
            case CfConst c: return c.Value;
            case CfRef r:
            {
                var (row, column) = ctx.Resolve(r);
                if (row < 0 || column < 0) return CfError.Value;
                if (r.Sheet is null) return ctx.CellAt(row, column);
                return ctx.Sheets?.Invoke(r.Sheet) is { } other ? other.CellAt(row, column) : CfError.Value;
            }
            case CfRange: return CfError.Value;
            case CfUnary u:
            {
                var v = Evaluate(u.Operand, ctx);
                if (v is CfError) return v;
                if (ToNumber(v) is not { } n) return CfError.Value;
                return u.Op == '-' ? -n : n / 100;
            }
            case CfBinary b: return Binary(b, ctx);
            case CfCall call: return Call(call, ctx);
        }
        return CfError.Value;
    }

    /// <summary>The truth value of a rule formula; errors and text are false, numbers are true unless zero.</summary>
    public static bool IsTrue(object value) => value switch
    {
        bool b => b,
        double d => d != 0,
        _ => false,
    };

    public static double? ToNumber(object value) => value switch
    {
        double d => d,
        bool b => b ? 1 : 0,
        CfBlank => 0,
        string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) && double.IsFinite(d) => d,
        _ => null,
    };

    public static string ToText(object value) => value switch
    {
        string s => s,
        double d => d.ToString("G15", CultureInfo.InvariantCulture),
        bool b => b ? "TRUE" : "FALSE",
        CfBlank => "",
        CfError e => e.Code,
        _ => "",
    };

    private static int Rank(object v) => v switch { double or CfBlank => 0, string => 1, bool => 2, _ => 3 };

    /// <summary>Excel's ordering: numbers before text before booleans; text compares without regard to case; blank equals 0 and "".</summary>
    public static int Compare(object a, object b)
    {
        if (a is CfBlank) a = b is string ? "" : b is bool ? false : 0.0;
        if (b is CfBlank) b = a is string ? "" : a is bool ? false : 0.0;
        int ra = Rank(a), rb = Rank(b);
        if (ra != rb) return ra.CompareTo(rb);
        return a switch
        {
            double x => x.CompareTo((double)b),
            string s => string.Compare(s, (string)b, StringComparison.OrdinalIgnoreCase),
            bool p => p.CompareTo((bool)b),
            _ => 0,
        };
    }

    private static object Binary(CfBinary b, CfContext ctx) => BinaryValues(b.Op, Evaluate(b.Left, ctx), Evaluate(b.Right, ctx));

    private static object BinaryValues(string op, object left, object right)
    {
        if (left is CfError) return left;
        if (right is CfError) return right;
        switch (op)
        {
            case "=": return Compare(left, right) == 0;
            case "<>": return Compare(left, right) != 0;
            case "<": return Compare(left, right) < 0;
            case "<=": return Compare(left, right) <= 0;
            case ">": return Compare(left, right) > 0;
            case ">=": return Compare(left, right) >= 0;
            case "&": return ToText(left) + ToText(right);
        }
        if (ToNumber(left) is not { } x || ToNumber(right) is not { } y) return CfError.Value;
        double result = op switch
        {
            "+" => x + y,
            "-" => x - y,
            "*" => x * y,
            "/" => y == 0 ? double.NaN : x / y,
            "^" => Math.Pow(x, y),
            _ => double.NaN,
        };
        if (double.IsNaN(result)) return op == "/" ? CfError.DivisionByZero : CfError.Value;
        return double.IsInfinity(result) ? CfError.Value : result;
    }

    /// <summary>The values an argument stands for: all cells of a range (inside the bounds) or the single value.</summary>
    private static IEnumerable<object> Values(CfExpr arg, CfContext ctx)
    {
        if (arg is CfCall && ReferenceOf(arg, ctx) is { } referenced) arg = referenced;
        if (arg is CfRange range)
        {
            var (r1, c1) = ctx.Resolve(range.From);
            var (r2, c2) = ctx.Resolve(range.To);
            var cellAt = ctx.CellAt;
            var bounds = (CellRange?)ctx.Bounds;
            if (range.From.Sheet is { } sheetName)
            {
                if (ctx.Sheets?.Invoke(sheetName) is not { } other)
                {
                    yield return CfError.Value;
                    yield break;
                }
                (cellAt, bounds) = (other.CellAt, other.Bounds);
            }
            if (bounds is not { } limit) yield break; // nothing on that sheet
            int firstRow = Math.Max(Math.Min(r1, r2), limit.FirstRow), lastRow = Math.Min(Math.Max(r1, r2), limit.LastRow);
            int firstColumn = Math.Max(Math.Min(c1, c2), limit.FirstColumn), lastColumn = Math.Min(Math.Max(c1, c2), limit.LastColumn);
            if (((long)lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) > MaxRangeCells)
            {
                yield return CfError.Value;
                yield break;
            }
            for (int r = firstRow; r <= lastRow; r++)
                for (int c = firstColumn; c <= lastColumn; c++)
                    yield return cellAt(r, c);
            yield break;
        }
        yield return Evaluate(arg, ctx);
    }

    private static object Call(CfCall call, CfContext ctx)
    {
        var args = call.Args;
        object Arg(int i) => i < args.Count ? Evaluate(args[i], ctx) : CfBlank.Instance;
        double? Num(int i) => ToNumber(Arg(i));
        bool Needs(int min, int max) => args.Count >= min && args.Count <= max;

        switch (call.Name)
        {
            case "TRUE": return true;
            case "FALSE": return false;
            case "AND":
            case "OR":
            case "XOR":
            {
                bool any = false, all = true;
                int trues = 0;
                foreach (var arg in args)
                    foreach (var v in Values(arg, ctx))
                    {
                        if (v is CfError) return v;
                        if (arg is CfRange && v is string or CfBlank) continue;
                        if (ToNumber(v) is not { } n) return CfError.Value;
                        any = true;
                        if (n != 0) trues++; else all = false;
                    }
                if (!any) return CfError.Value;
                return call.Name == "AND" ? all : call.Name == "OR" ? trues > 0 : trues % 2 == 1;
            }
            case "NOT":
            {
                if (!Needs(1, 1)) return CfError.Value;
                var v = Arg(0);
                return v is CfError ? v : ToNumber(v) is { } n ? n == 0 : CfError.Value;
            }
            case "IF":
            {
                if (!Needs(2, 3)) return CfError.Value;
                var condition = Arg(0);
                if (condition is CfError) return condition;
                if (ToNumber(condition) is not { } n) return CfError.Value;
                return n != 0 ? Arg(1) : args.Count > 2 ? Arg(2) : false;
            }
            case "IFERROR":
            {
                if (!Needs(2, 2)) return CfError.Value;
                var v = Arg(0);
                return v is CfError ? Arg(1) : v;
            }
            case "MOD":
            {
                if (!Needs(2, 2) || Num(0) is not { } x || Num(1) is not { } y) return CfError.Value;
                if (y == 0) return CfError.DivisionByZero;
                return x - y * Math.Floor(x / y);
            }
            case "ROW":
            case "COLUMN":
            {
                bool row = call.Name == "ROW";
                if (args.Count == 0) return (double)((row ? ctx.Row : ctx.Column) + 1);
                CfRef? target = args[0] as CfRef ?? (args[0] as CfRange)?.From;
                if (target is null) return CfError.Value;
                var (r, c) = ctx.Resolve(target);
                return (double)((row ? r : c) + 1);
            }
            case "ISBLANK": return Needs(1, 1) ? Arg(0) is CfBlank : CfError.Value;
            case "ISNUMBER": return Needs(1, 1) ? Arg(0) is double : CfError.Value;
            case "ISTEXT": return Needs(1, 1) ? Arg(0) is string : CfError.Value;
            case "ISERROR": return Needs(1, 1) ? Arg(0) is CfError : CfError.Value;
            case "ISEVEN":
            case "ISODD":
            {
                if (!Needs(1, 1) || Num(0) is not { } n) return CfError.Value;
                bool even = Math.Truncate(Math.Abs(n)) % 2 == 0;
                return call.Name == "ISEVEN" ? even : !even;
            }
            case "ABS": return Needs(1, 1) && Num(0) is { } a ? Math.Abs(a) : CfError.Value;
            case "INT": return Needs(1, 1) && Num(0) is { } i ? Math.Floor(i) : CfError.Value;
            case "ROUND":
            {
                if (!Needs(2, 2) || Num(0) is not { } x || Num(1) is not { } digits) return CfError.Value;
                return Math.Round(x, (int)Math.Clamp(digits, 0, 15), MidpointRounding.AwayFromZero);
            }
            case "LEN": return Needs(1, 1) && Arg(0) is var v0 and not CfError ? (double)ToText(v0).Length : CfError.Value;
            case "LEFT":
            case "RIGHT":
            {
                if (!Needs(1, 2)) return CfError.Value;
                var text = ToText(Arg(0));
                int count = args.Count > 1 ? (Num(1) is { } k && k >= 0 ? (int)Math.Min(k, text.Length) : -1) : 1;
                if (count < 0) return CfError.Value;
                return call.Name == "LEFT" ? text[..count] : text[^count..];
            }
            case "SEARCH":
            case "FIND":
            {
                if (!Needs(2, 3)) return CfError.Value;
                var needle = ToText(Arg(0));
                var haystack = ToText(Arg(1));
                int start = args.Count > 2 ? (Num(2) is { } s && s >= 1 && s <= haystack.Length + 1 ? (int)s - 1 : -1) : 0;
                if (start < 0) return CfError.Value;
                int at = haystack.IndexOf(needle, start, call.Name == "SEARCH" ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
                return at < 0 ? CfError.Value : (double)(at + 1);
            }
            case "SUM":
            case "MIN":
            case "MAX":
            case "AVERAGE":
            case "COUNT":
            case "COUNTA":
            {
                var numbers = new List<double>();
                int nonBlank = 0;
                foreach (var arg in args)
                    foreach (var v in Values(arg, ctx))
                    {
                        if (v is CfError && call.Name != "COUNT" && call.Name != "COUNTA") return v;
                        if (v is not CfBlank) nonBlank++;
                        if (v is double d) numbers.Add(d);
                        else if (arg is not CfRange && v is not CfError && call.Name != "COUNT" && v is not CfBlank && ToNumber(v) is { } direct) numbers.Add(direct);
                    }
                switch (call.Name)
                {
                    case "COUNT": return (double)numbers.Count;
                    case "COUNTA": return (double)nonBlank;
                    case "SUM": return numbers.Sum();
                    case "MIN": return numbers.Count == 0 ? 0.0 : numbers.Min();
                    case "MAX": return numbers.Count == 0 ? 0.0 : numbers.Max();
                    default: return numbers.Count == 0 ? CfError.DivisionByZero : numbers.Average();
                }
            }
            case "OFFSET":
            case "INDIRECT":
                return ReferenceOf(call, ctx) is { } pointed ? ValueAtReference(pointed, ctx) : CfError.Value;
            case "SUMPRODUCT":
            {
                if (args.Count == 0) return CfError.Value;
                var arrays = new List<object[,]>();
                foreach (var arg in args)
                {
                    var array = ArrayOf(arg, ctx);
                    if (array is null) return CfError.Value;
                    arrays.Add(array);
                }
                int rows = arrays[0].GetLength(0), columns = arrays[0].GetLength(1);
                if (arrays.Any(a => a.GetLength(0) != rows || a.GetLength(1) != columns)) return CfError.Value;
                double sum = 0;
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < columns; c++)
                    {
                        double product = 1;
                        foreach (var array in arrays)
                        {
                            if (array[r, c] is CfError error) return error;
                            product *= array[r, c] is double d ? d : 0; // text and booleans count as zero
                        }
                        sum += product;
                    }
                return sum;
            }
            case "INDEX":
            {
                if (!Needs(2, 3) || ArrayOf(args[0], ctx) is not { } grid) return CfError.Value;
                if (Num(1) is not { } rowNumber || (args.Count == 3 && Num(2) is null)) return CfError.Value;
                int row = (int)rowNumber, column = args.Count == 3 ? (int)Num(2)! : 0;
                int rows = grid.GetLength(0), columns = grid.GetLength(1);
                if (args.Count == 2 && (rows == 1 || columns == 1))
                {
                    // One row or one column: the single index counts along it.
                    int at = row - 1;
                    return at >= 0 && at < Math.Max(rows, columns) ? (rows == 1 ? grid[0, at] : grid[at, 0]) : CfError.Value;
                }
                if (args.Count == 2) column = 1;
                return row >= 1 && row <= rows && column >= 1 && column <= columns ? grid[row - 1, column - 1] : CfError.Value;
            }
            case "MATCH":
            {
                if (!Needs(2, 3) || ArrayOf(args[1], ctx) is not { } grid || (grid.GetLength(0) != 1 && grid.GetLength(1) != 1)) return CfError.Value;
                var key = Arg(0);
                if (key is CfError) return key;
                double type = args.Count == 3 ? Num(2) ?? 1 : 1;
                var list = grid.Cast<object>().ToList();
                return LookupPosition(key, list, type == 0 ? 0 : type > 0 ? 1 : -1) is { } position ? position + 1.0 : CfError.NotAvailable;
            }
            case "VLOOKUP":
            case "HLOOKUP":
            {
                if (!Needs(3, 4) || ArrayOf(args[1], ctx) is not { } table || Num(2) is not { } index) return CfError.Value;
                var key = Arg(0);
                if (key is CfError) return key;
                bool vertical = call.Name == "VLOOKUP";
                int length = vertical ? table.GetLength(0) : table.GetLength(1), width = vertical ? table.GetLength(1) : table.GetLength(0);
                if (index < 1 || index > width) return CfError.Value;
                bool exact = args.Count == 4 && Arg(3) is var approximate && !IsTrue(approximate is double or bool ? approximate : true);
                var firsts = Enumerable.Range(0, length).Select(i => vertical ? table[i, 0] : table[0, i]).ToList();
                if (LookupPosition(key, firsts, exact ? 0 : 1) is not { } found) return CfError.NotAvailable;
                return vertical ? table[found, (int)index - 1] : table[(int)index - 1, found];
            }
            case "COUNTIF":
            {
                if (!Needs(2, 2) || !(args[0] is CfRange || ReferenceOf(args[0], ctx) is CfRange)) return CfError.Value;
                var criteria = Arg(1);
                if (criteria is CfError) return criteria;
                var matches = Criteria(criteria);
                double count = 0;
                foreach (var v in Values(args[0], ctx))
                {
                    if (matches(v)) count++;
                }
                return count;
            }
        }
        return CfError.Value;
    }

    // ---- references and arrays

    /// <summary>
    /// The reference (absolute, as a <see cref="CfRef"/> or <see cref="CfRange"/>) that an expression stands for: a reference as it
    /// resolves for the current cell, or what OFFSET and INDIRECT point to. Null when the expression is no reference.
    /// </summary>
    private static CfExpr? ReferenceOf(CfExpr expr, CfContext ctx)
    {
        CfRef Absolute(CfRef r)
        {
            var (row, column) = ctx.Resolve(r);
            return new CfRef(row, true, column, true, r.Sheet);
        }
        switch (expr)
        {
            case CfRef r: return Absolute(r);
            case CfRange range: return new CfRange(Absolute(range.From), Absolute(range.To));
            case CfCall { Name: "OFFSET" } call when call.Args.Count is >= 3 and <= 5:
            {
                if (ReferenceOf(call.Args[0], ctx) is not { } baseReference) return null;
                double? Number(int i) => i < call.Args.Count ? ToNumber(Evaluate(call.Args[i], ctx)) : null;
                var (from, to) = baseReference is CfRange b ? (b.From, b.To) : ((CfRef)baseReference, (CfRef)baseReference);
                if (Number(1) is not { } rows || Number(2) is not { } columns) return null;
                int top = Math.Min(from.Row, to.Row), left = Math.Min(from.Column, to.Column);
                int height = call.Args.Count > 3 && Number(3) is { } h ? (int)h : Math.Abs(to.Row - from.Row) + 1;
                int width = call.Args.Count > 4 && Number(4) is { } w ? (int)w : Math.Abs(to.Column - from.Column) + 1;
                int row = top + (int)rows, column = left + (int)columns;
                if (height < 1 || width < 1 || row < 0 || column < 0 || row + height > CellReference.MaxRows || column + width > CellReference.MaxColumns) return null;
                var first = new CfRef(row, true, column, true, from.Sheet);
                return height == 1 && width == 1 ? first : new CfRange(first, new CfRef(row + height - 1, true, column + width - 1, true, from.Sheet));
            }
            case CfCall { Name: "INDIRECT" } call when call.Args.Count is 1 or 2:
            {
                if (Evaluate(call.Args[0], ctx) is not string text) return null;
                if (call.Args.Count == 2 && !IsTrue(Evaluate(call.Args[1], ctx))) return null; // R1C1 text
                CfExpr parsed;
                try { parsed = CfParser.Parse(text); }
                catch (NotSupportedException) { return null; }
                // Text points at fixed places: it does not move with the cell the formula is evaluated for.
                static CfRef Fixed(CfRef r) => new(r.Row, true, r.Column, true, r.Sheet);
                return parsed switch { CfRef r => Fixed(r), CfRange range => new CfRange(Fixed(range.From), Fixed(range.To)), _ => null };
            }
            default: return null;
        }
    }

    /// <summary>The value at an absolute reference (the top-left cell of a range).</summary>
    private static object ValueAtReference(CfExpr reference, CfContext ctx)
    {
        var cell = reference is CfRange range ? range.From : (CfRef)reference;
        var (row, column) = (cell.Row, cell.Column);
        if (reference is CfRange r) (row, column) = (Math.Min(r.From.Row, r.To.Row), Math.Min(r.From.Column, r.To.Column));
        if (cell.Sheet is null) return ctx.CellAt(row, column);
        return ctx.Sheets?.Invoke(cell.Sheet) is { } other ? other.CellAt(row, column) : CfError.Value;
    }

    /// <summary>
    /// The values an expression stands for as a grid: a range or the result of OFFSET/INDIRECT, a single value, or the cell-by-cell
    /// result of arithmetic and comparisons over those (a scalar is repeated across the grid). Null when the sizes do not match.
    /// </summary>
    private static object[,]? ArrayOf(CfExpr expr, CfContext ctx)
    {
        if (expr is CfCall && ReferenceOf(expr, ctx) is { } referenced) expr = referenced;
        switch (expr)
        {
            case CfRef or CfRange:
            {
                var (cellAt, sheetName) = (ctx.CellAt, (expr as CfRef ?? ((CfRange)expr).From).Sheet);
                if (sheetName is not null)
                {
                    if (ctx.Sheets?.Invoke(sheetName) is not { } other) return new object[,] { { CfError.Value } };
                    cellAt = other.CellAt;
                }
                var (from, to) = expr is CfRange range ? (ctx.Resolve(range.From), ctx.Resolve(range.To)) : (ctx.Resolve((CfRef)expr), ctx.Resolve((CfRef)expr));
                int r0 = Math.Min(from.Row, to.Row), r1 = Math.Max(from.Row, to.Row), c0 = Math.Min(from.Column, to.Column), c1 = Math.Max(from.Column, to.Column);
                if (r0 < 0 || c0 < 0) return null;
                if (((long)r1 - r0 + 1) * (c1 - c0 + 1) > MaxRangeCells)
                {
                    // A whole column or row: only the part the sheet can hold values in.
                    var limit = sheetName is null ? ctx.Bounds : ctx.Sheets?.Invoke(sheetName)?.Bounds;
                    if (limit is not { } bounds) return null;
                    (r0, r1, c0, c1) = (Math.Max(r0, bounds.FirstRow), Math.Min(r1, bounds.LastRow), Math.Max(c0, bounds.FirstColumn), Math.Min(c1, bounds.LastColumn));
                    if (r0 > r1 || c0 > c1 || ((long)r1 - r0 + 1) * (c1 - c0 + 1) > MaxRangeCells) return null;
                }
                var grid = new object[r1 - r0 + 1, c1 - c0 + 1];
                for (int r = r0; r <= r1; r++)
                    for (int c = c0; c <= c1; c++) grid[r - r0, c - c0] = cellAt(r, c);
                return grid;
            }
            case CfUnary u:
            {
                if (ArrayOf(u.Operand, ctx) is not { } inner) return null;
                var result = new object[inner.GetLength(0), inner.GetLength(1)];
                for (int r = 0; r < inner.GetLength(0); r++)
                    for (int c = 0; c < inner.GetLength(1); c++)
                        result[r, c] = inner[r, c] is CfError ? inner[r, c] : ToNumber(inner[r, c]) is { } n ? u.Op == '-' ? -n : n / 100 : CfError.Value;
                return result;
            }
            case CfBinary b:
            {
                if (ArrayOf(b.Left, ctx) is not { } left || ArrayOf(b.Right, ctx) is not { } right) return null;
                int rows = Math.Max(left.GetLength(0), right.GetLength(0)), columns = Math.Max(left.GetLength(1), right.GetLength(1));
                bool Fits(object[,] a) => (a.GetLength(0) == rows || a.GetLength(0) == 1) && (a.GetLength(1) == columns || a.GetLength(1) == 1);
                if (!Fits(left) || !Fits(right)) return null;
                var result = new object[rows, columns];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < columns; c++)
                        result[r, c] = BinaryValues(b.Op, left[left.GetLength(0) == 1 ? 0 : r, left.GetLength(1) == 1 ? 0 : c], right[right.GetLength(0) == 1 ? 0 : r, right.GetLength(1) == 1 ? 0 : c]);
                return result;
            }
            default:
                return new[,] { { Evaluate(expr, ctx) } };
        }
    }

    /// <summary>
    /// The position of <paramref name="key"/> in <paramref name="list"/>: 0 finds the first equal item; 1 the last item not greater (ascending
    /// data); -1 the last item not less (descending data). Only items of the same kind (number or text) are compared. Null when absent.
    /// </summary>
    private static int? LookupPosition(object key, List<object> list, int mode)
    {
        if (key is CfBlank) return null;
        int? found = null;
        for (int i = 0; i < list.Count; i++)
        {
            var item = list[i];
            if (item is CfError || item is CfBlank || Rank(item) != Rank(key)) continue;
            int cmp = Compare(item, key);
            if (mode == 0) { if (cmp == 0) return i; }
            else if (mode > 0 ? cmp <= 0 : cmp >= 0) found = i;
            else if (found is not null) break; // sorted data: nothing further can match
        }
        return found;
    }

    /// <summary>A COUNTIF criterion: a value, or text such as <c>"&gt;5"</c> / <c>"&lt;&gt;x"</c>; text compares without regard to case.</summary>
    private static Func<object, bool> Criteria(object criteria)
    {
        if (criteria is string text)
        {
            foreach (var op in new[] { "<=", ">=", "<>", "<", ">", "=" })
            {
                if (!text.StartsWith(op, StringComparison.Ordinal)) continue;
                var rest = text[op.Length..];
                object operand = double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : rest;
                return v =>
                {
                    if (v is CfBlank || v is CfError) return op == "<>" && rest.Length > 0 && v is CfBlank;
                    // Numbers only match numbers, text only text.
                    if (Rank(v) != Rank(operand)) return op == "<>";
                    int cmp = Compare(v, operand);
                    return op switch { "<=" => cmp <= 0, ">=" => cmp >= 0, "<>" => cmp != 0, "<" => cmp < 0, ">" => cmp > 0, _ => cmp == 0 };
                };
            }
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var asNumber))
                return v => v is double d && d == asNumber || v is string s && string.Equals(s, text, StringComparison.OrdinalIgnoreCase);
            return v => v is string s && string.Equals(s, text, StringComparison.OrdinalIgnoreCase);
        }
        return v => v is not CfBlank && v is not CfError && Rank(v) == Rank(criteria) && Compare(v, criteria) == 0;
    }
}
