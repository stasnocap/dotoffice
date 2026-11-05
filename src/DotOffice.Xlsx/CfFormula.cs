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
internal sealed record CfRef(int Row, bool RowAbsolute, int Column, bool ColumnAbsolute) : CfExpr;
internal sealed record CfRange(CfRef From, CfRef To) : CfExpr;
internal sealed record CfUnary(char Op, CfExpr Operand) : CfExpr;
internal sealed record CfBinary(string Op, CfExpr Left, CfExpr Right) : CfExpr;
internal sealed record CfCall(string Name, IReadOnlyList<CfExpr> Args) : CfExpr;

/// <summary>
/// Parser for the formulas that conditional-formatting rules use: arithmetic, comparison, text concatenation, cell references
/// and ranges on the same sheet, and a fixed set of functions. Anything else throws <see cref="NotSupportedException"/>.
/// </summary>
internal sealed class CfParser
{
    public static readonly HashSet<string> Functions = new(StringComparer.Ordinal)
    {
        "AND", "OR", "NOT", "XOR", "IF", "IFERROR", "MOD", "ROW", "COLUMN", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISEVEN", "ISODD",
        "ABS", "INT", "ROUND", "LEN", "LEFT", "RIGHT", "SEARCH", "FIND", "SUM", "MIN", "MAX", "AVERAGE", "COUNT", "COUNTA", "COUNTIF", "TRUE", "FALSE",
    };

    private readonly string _text;
    private int _i;

    private CfParser(string text) => _text = text;

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
            if (_i < _text.Length && _text[_i] == '!') throw new NotSupportedException("References to other sheets are not supported.");
            if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return new CfConst(true);
            if (word.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return new CfConst(false);
            var from = Reference(word) ?? throw new NotSupportedException($"'{word}' is not a cell reference or a supported name.");
            if (_i < _text.Length && _text[_i] == ':')
            {
                _i++;
                int s2 = _i;
                while (_i < _text.Length && (char.IsAsciiLetterOrDigit(_text[_i]) || _text[_i] == '$')) _i++;
                var to = Reference(_text[s2.._i]) ?? throw new NotSupportedException("Unsupported range.");
                return new CfRange(from, to);
            }
            return from;
        }
        throw new NotSupportedException($"Unexpected '{c}' in formula.");
    }

    private static CfRef? Reference(string word)
    {
        if (!CellReference.TryParse(word, out var row, out var column)) return null;
        return new CfRef(row, HasRowDollar(word), column, word[0] == '$');
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
internal sealed class CfContext(Func<int, int, object> cellAt, int originRow, int originColumn, int row, int column, CellRange bounds)
{
    public Func<int, int, object> CellAt { get; } = cellAt;

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
                return row < 0 || column < 0 ? CfError.Value : ctx.CellAt(row, column);
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

    private static object Binary(CfBinary b, CfContext ctx)
    {
        var left = Evaluate(b.Left, ctx);
        var right = Evaluate(b.Right, ctx);
        if (left is CfError) return left;
        if (right is CfError) return right;
        switch (b.Op)
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
        double result = b.Op switch
        {
            "+" => x + y,
            "-" => x - y,
            "*" => x * y,
            "/" => y == 0 ? double.NaN : x / y,
            "^" => Math.Pow(x, y),
            _ => double.NaN,
        };
        if (double.IsNaN(result)) return b.Op == "/" ? CfError.DivisionByZero : CfError.Value;
        return double.IsInfinity(result) ? CfError.Value : result;
    }

    /// <summary>The values an argument stands for: all cells of a range (inside the bounds) or the single value.</summary>
    private static IEnumerable<object> Values(CfExpr arg, CfContext ctx)
    {
        if (arg is CfRange range)
        {
            var (r1, c1) = ctx.Resolve(range.From);
            var (r2, c2) = ctx.Resolve(range.To);
            int firstRow = Math.Max(Math.Min(r1, r2), ctx.Bounds.FirstRow), lastRow = Math.Min(Math.Max(r1, r2), ctx.Bounds.LastRow);
            int firstColumn = Math.Max(Math.Min(c1, c2), ctx.Bounds.FirstColumn), lastColumn = Math.Min(Math.Max(c1, c2), ctx.Bounds.LastColumn);
            if (((long)lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) > MaxRangeCells)
            {
                yield return CfError.Value;
                yield break;
            }
            for (int r = firstRow; r <= lastRow; r++)
                for (int c = firstColumn; c <= lastColumn; c++)
                    yield return ctx.CellAt(r, c);
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
            case "COUNTIF":
            {
                if (!Needs(2, 2) || args[0] is not CfRange) return CfError.Value;
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
