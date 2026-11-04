using System.Globalization;
using System.Text;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>A formatted cell value plus the color requested by a <c>[Red]</c>-style section (null = keep the font color).</summary>
internal readonly record struct FormattedValue(string Text, Color? Color);

/// <summary>A parsed Excel number-format code (<c>pos;neg;zero;text</c>) that can format numbers, dates and text.</summary>
internal sealed class NumberFormat
{
    private enum Kind { Literal, Digit, Point, Comma, Percent, Exponent, Text, General, Year, Month, Day, Hour, Minute, Second, Elapsed, AmPm, SubSecond }

    private readonly record struct Token(Kind Kind, string Text = "", int Count = 0);

    private sealed class Section
    {
        public List<Token> Tokens = [];
        public Color? Color;
        public (char Op, double Value)? Condition;
        public bool IsDate;
        public bool HasText;
    }

    private readonly List<Section> _sections = [];

    private static readonly Dictionary<string, NumberFormat> Cache = new();

    /// <summary>Parses (and caches) a format code. Never throws; unparseable codes degrade to General.</summary>
    public static NumberFormat Parse(string code)
    {
        lock (Cache)
        {
            if (!Cache.TryGetValue(code, out var format))
            {
                format = new NumberFormat(code);
                if (Cache.Count < 4096) Cache[code] = format;
            }
            return format;
        }
    }

    /// <summary>The format code of a built-in number format id, or null when the id is unknown (treated as General).</summary>
    public static string? BuiltIn(int id) => id switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);\\(\"$\"#,##0\\)",
        6 => "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)",
        7 => "\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)",
        8 => "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        41 => "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };

    private static readonly string[] MonthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    private static readonly string[] DayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

    private static readonly Dictionary<string, Color> ColorNames = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = new(0, 0, 0), ["white"] = new(255, 255, 255), ["red"] = new(255, 0, 0), ["green"] = new(0, 255, 0),
        ["blue"] = new(0, 0, 255), ["yellow"] = new(255, 255, 0), ["magenta"] = new(255, 0, 255), ["cyan"] = new(0, 255, 255),
    };

    // ------------------------------------------------------------------ parsing

    private NumberFormat(string code)
    {
        foreach (var part in SplitSections(code)) _sections.Add(ParseSection(part));
        if (_sections.Count == 0) _sections.Add(new Section { Tokens = [new Token(Kind.General)] });
    }

    private static List<string> SplitSections(string code)
    {
        var parts = new List<string>();
        var sb = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (c == '"') quoted = !quoted;
            else if (!quoted && c == '\\' && i + 1 < code.Length) { sb.Append(c).Append(code[++i]); continue; }
            else if (!quoted && c == ';') { parts.Add(sb.ToString()); sb.Clear(); continue; }
            sb.Append(c);
        }
        parts.Add(sb.ToString());
        return parts;
    }

    private static Section ParseSection(string text)
    {
        var s = new Section();
        var tokens = s.Tokens;
        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            switch (c)
            {
                case '"':
                {
                    int end = text.IndexOf('"', i + 1);
                    if (end < 0) end = text.Length;
                    AddLiteral(tokens, text.Substring(i + 1, end - i - 1));
                    i = end + 1;
                    continue;
                }
                case '\\':
                    if (i + 1 < text.Length) AddLiteral(tokens, text[i + 1].ToString());
                    i += 2;
                    continue;
                case '_':
                    if (i + 1 < text.Length) AddLiteral(tokens, " ");
                    i += 2;
                    continue;
                case '*':
                    i += 2; // repeat-fill is ignored
                    continue;
                case '[':
                {
                    int end = text.IndexOf(']', i);
                    if (end < 0) { i = text.Length; continue; }
                    ParseBracket(s, text.Substring(i + 1, end - i - 1));
                    i = end + 1;
                    continue;
                }
                case '0' or '#' or '?':
                    tokens.Add(new Token(Kind.Digit, c.ToString()));
                    break;
                case '.':
                    // ".0" right after a seconds token is a fraction of a second.
                    if (LastDate(tokens) is Kind.Second && i + 1 < text.Length && text[i + 1] == '0')
                    {
                        int n = 0;
                        while (i + 1 + n < text.Length && text[i + 1 + n] == '0') n++;
                        tokens.Add(new Token(Kind.SubSecond, Count: n));
                        s.IsDate = true;
                        i += 1 + n;
                        continue;
                    }
                    tokens.Add(new Token(Kind.Point));
                    break;
                case ',':
                    tokens.Add(new Token(Kind.Comma));
                    break;
                case '%':
                    tokens.Add(new Token(Kind.Percent));
                    break;
                case '@':
                    tokens.Add(new Token(Kind.Text));
                    s.HasText = true;
                    break;
                case 'E' or 'e' when i + 1 < text.Length && text[i + 1] is '+' or '-':
                {
                    int n = 0;
                    while (i + 2 + n < text.Length && text[i + 2 + n] == '0') n++;
                    tokens.Add(new Token(Kind.Exponent, text[i + 1].ToString(), Math.Max(n, 1)));
                    i += 2 + n;
                    continue;
                }
                case 'y' or 'Y' or 'd' or 'D' or 'm' or 'M' or 'h' or 'H' or 's' or 'S':
                {
                    char lower = char.ToLowerInvariant(c);
                    int n = 0;
                    while (i + n < text.Length && char.ToLowerInvariant(text[i + n]) == lower) n++;
                    var kind = lower switch { 'y' => Kind.Year, 'm' => Kind.Month, 'd' => Kind.Day, 'h' => Kind.Hour, _ => Kind.Second };
                    tokens.Add(new Token(kind, Count: n));
                    s.IsDate = true;
                    i += n;
                    continue;
                }
                case 'a' or 'A' or 'p' or 'P':
                {
                    var rest = text.AsSpan(i);
                    if (rest.StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase))
                    {
                        tokens.Add(new Token(Kind.AmPm, text.Substring(i, 5), 2));
                        s.IsDate = true;
                        i += 5;
                        continue;
                    }
                    if (rest.StartsWith("A/P", StringComparison.OrdinalIgnoreCase))
                    {
                        tokens.Add(new Token(Kind.AmPm, text.Substring(i, 3), 1));
                        s.IsDate = true;
                        i += 3;
                        continue;
                    }
                    AddLiteral(tokens, c.ToString());
                    break;
                }
                case 'G' or 'g' when text.AsSpan(i).StartsWith("General", StringComparison.OrdinalIgnoreCase):
                    tokens.Add(new Token(Kind.General));
                    i += 7;
                    continue;
                default:
                    AddLiteral(tokens, c.ToString());
                    break;
            }
            i++;
        }
        return s;
    }

    private static Kind? LastDate(List<Token> tokens)
    {
        for (int i = tokens.Count - 1; i >= 0; i--)
            if (tokens[i].Kind != Kind.Literal) return tokens[i].Kind;
        return null;
    }

    private static void AddLiteral(List<Token> tokens, string text)
    {
        if (tokens.Count > 0 && tokens[^1].Kind == Kind.Literal) tokens[^1] = tokens[^1] with { Text = tokens[^1].Text + text };
        else tokens.Add(new Token(Kind.Literal, text));
    }

    private static void ParseBracket(Section s, string content)
    {
        if (content.Length == 0) return;
        if (ColorNames.TryGetValue(content, out var color)) { s.Color = color; return; }
        if (content.StartsWith("color", StringComparison.OrdinalIgnoreCase)) return; // indexed palette colors are not supported
        if (content[0] is '<' or '>' or '=')
        {
            int pos = 1;
            string op = content[0].ToString();
            if (content.Length > 1 && content[1] is '=' or '>')
            {
                op += content[1];
                pos = 2;
            }
            if (double.TryParse(content.AsSpan(pos), NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                s.Condition = (op switch { "<" => '<', ">" => '>', "=" => '=', "<=" => 'l', ">=" => 'g', _ => 'n' }, v);
            return;
        }
        if (content[0] == '$')
        {
            // [$€-407]: currency symbol before the locale suffix.
            int dash = content.IndexOf('-');
            string symbol = dash < 0 ? content[1..] : content[1..dash];
            if (symbol.Length > 0) AddLiteral(s.Tokens, symbol);
            return;
        }
        char first = char.ToLowerInvariant(content[0]);
        if (first is 'h' or 'm' or 's' && content.All(ch => char.ToLowerInvariant(ch) == first))
        {
            s.Tokens.Add(new Token(Kind.Elapsed, first.ToString(), content.Length));
            s.IsDate = true;
        }
    }

    // ------------------------------------------------------------------ formatting

    /// <summary>Formats a number with this code.</summary>
    public FormattedValue FormatNumber(double value, bool date1904)
    {
        if (!double.IsFinite(value)) return new FormattedValue("#NUM!", null);
        int index = PickSection(value, out bool showSign);
        var section = _sections[index];
        double v = showSign ? value : Math.Abs(value);
        if (section.HasText && !section.IsDate && !HasNumeric(section)) return new FormattedValue(string.Concat(section.Tokens.Select(t => t.Kind == Kind.Literal ? t.Text : "")), section.Color);

        string text = section.IsDate ? FormatDate(section, value, date1904) : FormatDecimal(section, v);
        return new FormattedValue(text, section.Color);
    }

    /// <summary>Formats a text cell value: only a text section (<c>@</c>) changes it.</summary>
    public FormattedValue FormatText(string value)
    {
        Section? section = null;
        if (_sections.Count >= 4) section = _sections[3];
        else if (_sections.Count == 1 && _sections[0].HasText) section = _sections[0];
        if (section is null) return new FormattedValue(value, null);
        var sb = new StringBuilder();
        foreach (var t in section.Tokens)
            if (t.Kind == Kind.Text) sb.Append(value);
            else if (t.Kind == Kind.Literal) sb.Append(t.Text);
        return new FormattedValue(sb.ToString(), section.Color);
    }

    private static bool HasNumeric(Section s) => s.Tokens.Any(t => t.Kind is Kind.Digit or Kind.General);

    private int PickSection(double value, out bool showSign)
    {
        int n = Math.Min(_sections.Count, 3);
        if (_sections.Count == 4 || n > 0)
        {
            // Any explicit condition → evaluate in order.
            bool conditional = false;
            for (int i = 0; i < n; i++) conditional |= _sections[i].Condition is not null;
            if (conditional)
            {
                for (int i = 0; i < n; i++)
                {
                    var cond = _sections[i].Condition ?? (i == 0 ? (n > 1 ? ('>', 0) : ('g', double.NegativeInfinity)) : i == 1 ? ('<', 0) : ('g', double.NegativeInfinity));
                    if (Matches(cond.Item1, cond.Item2, value))
                    {
                        showSign = value < 0 && !(i == 1 && _sections[1].Condition is null);
                        return i;
                    }
                }
                showSign = value < 0;
                return n - 1;
            }
        }
        if (n == 1) { showSign = value < 0; return 0; }
        if (n == 2)
        {
            if (value < 0) { showSign = false; return 1; }
            showSign = false;
            return 0;
        }
        if (value > 0) { showSign = false; return 0; }
        if (value < 0) { showSign = false; return 1; }
        showSign = false;
        return 2;
    }

    private static bool Matches(char op, double limit, double v) => op switch
    {
        '<' => v < limit,
        '>' => v > limit,
        '=' => v == limit,
        'l' => v <= limit,
        'g' => v >= limit,
        _ => v != limit,
    };

    // ---- numbers

    private static string FormatDecimal(Section s, double value)
    {
        var tokens = s.Tokens;
        bool negative = value < 0;
        double abs = Math.Abs(value);

        // Layout facts.
        int pointIndex = tokens.FindIndex(t => t.Kind == Kind.Point);
        int intPlaceholders = 0, fracPlaceholders = 0, lastDigit = -1, firstDigit = -1;
        for (int i = 0; i < tokens.Count; i++)
        {
            if (tokens[i].Kind != Kind.Digit) continue;
            if (firstDigit < 0) firstDigit = i;
            lastDigit = i;
            if (pointIndex >= 0 && i > pointIndex) fracPlaceholders++;
            else intPlaceholders++;
        }
        bool hasGeneral = tokens.Any(t => t.Kind == Kind.General);
        if (firstDigit < 0 && !tokens.Any(t => t.Kind == Kind.Exponent))
        {
            if (hasGeneral)
                return Assemble(tokens, general: NumberFormatter.General(negative ? -abs : abs), digits: null);
            // Literal-only section (e.g. `"zero"`): just the literals (percent still prints).
            return Assemble(tokens, general: null, digits: null);
        }

        int percents = tokens.Count(t => t.Kind == Kind.Percent);
        for (int i = 0; i < percents; i++) abs *= 100;

        // Scaling commas: right after the last digit placeholder of the integer part (or directly before the point).
        int scale = 0;
        int intEnd = pointIndex >= 0 ? pointIndex : tokens.Count;
        int lastIntDigit = -1;
        for (int i = 0; i < intEnd; i++) if (tokens[i].Kind == Kind.Digit) lastIntDigit = i;
        // Commas directly after the last digit placeholder (before the point, or at the very end) scale by 1000 each.
        for (int i = lastIntDigit + 1; lastIntDigit >= 0 && i < intEnd && tokens[i].Kind == Kind.Comma; i++) scale++;
        if (pointIndex >= 0 && lastDigit > pointIndex)
            for (int i = lastDigit + 1; i < tokens.Count && tokens[i].Kind == Kind.Comma; i++) scale++;
        bool thousands = false;
        for (int i = firstDigit; i < lastIntDigit; i++) thousands |= tokens[i].Kind == Kind.Comma;
        abs /= Math.Pow(1000, scale);

        int expIndex = tokens.FindIndex(t => t.Kind == Kind.Exponent);
        int exponent = 0;
        if (expIndex >= 0)
        {
            if (abs != 0)
            {
                exponent = (int)Math.Floor(Math.Log10(abs)) - (Math.Max(intPlaceholders, 1) - 1);
                abs /= Math.Pow(10, exponent);
            }
        }

        // Round to the fractional digit count.
        string intDigits, fracDigits;
        RoundSplit(abs, fracPlaceholders, out intDigits, out fracDigits);
        if (expIndex >= 0 && abs != 0 && intDigits.Length > Math.Max(intPlaceholders, 1))
        {
            // Rounding carried (9.99 → 10.0): renormalise.
            exponent++;
            abs /= 10;
            RoundSplit(abs, fracPlaceholders, out intDigits, out fracDigits);
        }
        if (intDigits == "0") intDigits = "";

        bool isZero = intDigits.Length == 0 && fracDigits.All(ch => ch == '0');
        string signPrefix = negative && !isZero ? "-" : "";
        var digits = new DigitState(intDigits, fracDigits, thousands, intPlaceholders, exponent, expIndex >= 0 ? tokens[expIndex] : default);
        return signPrefix + Assemble(tokens, general: null, digits);
    }

    private sealed record DigitState(string IntDigits, string FracDigits, bool Thousands, int IntPlaceholders, int Exponent, Token Exp);

    private static void RoundSplit(double value, int fractionDigits, out string intPart, out string fracPart)
    {
        fractionDigits = Math.Min(fractionDigits, 20);
        string text;
        if (value < 7.9e27 && value >= 0)
            text = Math.Round((decimal)value, Math.Min(fractionDigits, 28), MidpointRounding.AwayFromZero).ToString("F" + fractionDigits, CultureInfo.InvariantCulture);
        else
            text = value.ToString("F" + fractionDigits, CultureInfo.InvariantCulture);
        int dot = text.IndexOf('.');
        intPart = dot < 0 ? text : text[..dot];
        fracPart = dot < 0 ? "" : text[(dot + 1)..];
        intPart = intPart.TrimStart('0');
        if (intPart.Length == 0) intPart = "0";
    }

    private static string Assemble(List<Token> tokens, string? general, DigitState? digits)
    {
        var sb = new StringBuilder();
        if (digits is null)
        {
            foreach (var t in tokens)
            {
                if (t.Kind == Kind.Literal) sb.Append(t.Text);
                else if (t.Kind == Kind.Percent) sb.Append('%');
                else if (t.Kind == Kind.General && general is not null) sb.Append(general);
            }
            return sb.ToString();
        }

        string intStr = digits.IntDigits;
        int len = intStr.Length;
        int placeholdersLeft = digits.IntPlaceholders;
        int pointIndex = tokens.FindIndex(t => t.Kind == Kind.Point);
        int fracIndex = 0;

        // Trailing '#'/'?' zeros in the fractional part are dropped / blanked.
        var fracPlaceholderIdx = new List<int>();
        for (int i = 0; i < tokens.Count; i++)
            if (pointIndex >= 0 && i > pointIndex && tokens[i].Kind == Kind.Digit) fracPlaceholderIdx.Add(i);
        string frac = digits.FracDigits.PadRight(fracPlaceholderIdx.Count, '0');
        int keep = frac.Length;
        while (keep > 0 && fracPlaceholderIdx.Count >= keep && tokens[fracPlaceholderIdx[keep - 1]].Text != "0" && frac[keep - 1] == '0') keep--;

        bool extraEmitted = false;
        for (int i = 0; i < tokens.Count; i++)
        {
            var t = tokens[i];
            switch (t.Kind)
            {
                case Kind.Literal:
                    sb.Append(t.Text);
                    break;
                case Kind.Percent:
                    sb.Append('%');
                    break;
                case Kind.Point:
                    sb.Append('.');
                    break;
                case Kind.Exponent:
                    sb.Append('E');
                    if (digits.Exponent < 0) sb.Append('-');
                    else if (t.Text == "+") sb.Append('+');
                    sb.Append(Math.Abs(digits.Exponent).ToString(CultureInfo.InvariantCulture).PadLeft(t.Count, '0'));
                    break;
                case Kind.Digit when pointIndex >= 0 && i > pointIndex:
                {
                    int k = fracIndex++;
                    if (k < keep) sb.Append(frac[k]);
                    else if (t.Text == "?") sb.Append(' ');
                    break;
                }
                case Kind.Digit:
                {
                    if (!extraEmitted)
                    {
                        extraEmitted = true;
                        int extra = len - digits.IntPlaceholders;
                        for (int e = 0; e < extra; e++)
                        {
                            sb.Append(intStr[e]);
                            int pos = len - 1 - e;
                            if (digits.Thousands && pos > 0 && pos % 3 == 0) sb.Append(',');
                        }
                    }
                    int position = placeholdersLeft - 1; // digit position from the right
                    placeholdersLeft--;
                    if (position < len)
                    {
                        sb.Append(intStr[len - 1 - position]);
                        if (digits.Thousands && position > 0 && position % 3 == 0) sb.Append(',');
                    }
                    else if (t.Text == "0")
                    {
                        sb.Append('0');
                        if (digits.Thousands && position > 0 && position % 3 == 0) sb.Append(',');
                    }
                    else if (t.Text == "?") sb.Append(' ');
                    break;
                }
                case Kind.General when general is not null:
                    sb.Append(general);
                    break;
            }
        }
        return sb.ToString();
    }

    // ---- dates

    private static string FormatDate(Section s, double serial, bool date1904)
    {
        if (serial < 0) return "########";
        bool subSecond = s.Tokens.Any(t => t.Kind == Kind.SubSecond);
        int subDigits = subSecond ? s.Tokens.First(t => t.Kind == Kind.SubSecond).Count : 0;
        bool twelveHour = s.Tokens.Any(t => t.Kind == Kind.AmPm);

        double day = Math.Floor(serial);
        double frac = serial - day;
        long totalMs = subSecond ? (long)Math.Round(frac * 86_400_000d) : (long)Math.Round(frac * 86_400d) * 1000;
        if (subSecond && subDigits < 3)
        {
            long unit = (long)Math.Pow(10, 3 - subDigits);
            totalMs = (long)Math.Round(totalMs / (double)unit) * unit;
        }
        if (totalMs >= 86_400_000) { totalMs -= 86_400_000; day++; }

        int year, month, dom, weekday;
        if (!date1904 && day < 61)
        {
            // Lotus bug: the 1900 date system pretends 1900 was a leap year.
            weekday = (int)((day + 6) % 7);
            if (day == 60) { year = 1900; month = 2; dom = 29; }
            else if (day == 0) { year = 1900; month = 1; dom = 0; }
            else
            {
                var d = new DateTime(1899, 12, 31).AddDays(day);
                year = d.Year; month = d.Month; dom = d.Day;
            }
        }
        else
        {
            var baseDate = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
            if (day > 2_958_465) return "########";
            var d = baseDate.AddDays(day);
            year = d.Year; month = d.Month; dom = d.Day; weekday = (int)d.DayOfWeek;
        }

        long totalSeconds = (long)day * 86_400 + totalMs / 1000;
        int hours = (int)(totalMs / 3_600_000);
        int minutes = (int)(totalMs / 60_000 % 60);
        int seconds = (int)(totalMs / 1000 % 60);
        int millis = (int)(totalMs % 1000);

        var tokens = s.Tokens;
        var sb = new StringBuilder();
        for (int i = 0; i < tokens.Count; i++)
        {
            var t = tokens[i];
            switch (t.Kind)
            {
                case Kind.Literal:
                    sb.Append(t.Text);
                    break;
                case Kind.Comma:
                    sb.Append(',');
                    break;
                case Kind.Point:
                    sb.Append('.');
                    break;
                case Kind.Year:
                    if (t.Count <= 2) sb.Append((year % 100).ToString("00", CultureInfo.InvariantCulture));
                    else sb.Append(year.ToString("0000", CultureInfo.InvariantCulture));
                    break;
                case Kind.Month when IsMinute(tokens, i):
                    sb.Append(t.Count >= 2 ? minutes.ToString("00", CultureInfo.InvariantCulture) : minutes.ToString(CultureInfo.InvariantCulture));
                    break;
                case Kind.Month:
                    sb.Append(t.Count switch
                    {
                        1 => month.ToString(CultureInfo.InvariantCulture),
                        2 => month.ToString("00", CultureInfo.InvariantCulture),
                        3 => MonthNames[month - 1][..3],
                        4 => MonthNames[month - 1],
                        _ => MonthNames[month - 1][..1],
                    });
                    break;
                case Kind.Day:
                    sb.Append(t.Count switch
                    {
                        1 => dom.ToString(CultureInfo.InvariantCulture),
                        2 => dom.ToString("00", CultureInfo.InvariantCulture),
                        3 => DayNames[weekday][..3],
                        _ => DayNames[weekday],
                    });
                    break;
                case Kind.Hour:
                {
                    int h = twelveHour ? (hours % 12 == 0 ? 12 : hours % 12) : hours;
                    sb.Append(t.Count >= 2 ? h.ToString("00", CultureInfo.InvariantCulture) : h.ToString(CultureInfo.InvariantCulture));
                    break;
                }
                case Kind.Second:
                    sb.Append(t.Count >= 2 ? seconds.ToString("00", CultureInfo.InvariantCulture) : seconds.ToString(CultureInfo.InvariantCulture));
                    break;
                case Kind.SubSecond:
                    sb.Append('.').Append(millis.ToString("000", CultureInfo.InvariantCulture)[..Math.Min(t.Count, 3)].PadRight(t.Count, '0'));
                    break;
                case Kind.Elapsed:
                {
                    long value = t.Text switch { "h" => totalSeconds / 3600, "m" => totalSeconds / 60, _ => totalSeconds };
                    sb.Append(value.ToString(CultureInfo.InvariantCulture).PadLeft(t.Count, '0'));
                    break;
                }
                case Kind.AmPm:
                {
                    bool pm = hours >= 12;
                    string text = t.Count == 2 ? (pm ? "PM" : "AM") : (pm ? "P" : "A");
                    sb.Append(char.IsLower(t.Text[0]) ? text.ToLowerInvariant() : text);
                    break;
                }
            }
        }
        return sb.ToString();
    }

    /// <summary>"m" means minutes right after an hour token or right before a seconds token; otherwise it is the month.</summary>
    private static bool IsMinute(List<Token> tokens, int index)
    {
        for (int i = index - 1; i >= 0; i--)
        {
            if (tokens[i].Kind == Kind.Literal) continue;
            if (tokens[i].Kind == Kind.Hour || (tokens[i].Kind == Kind.Elapsed && tokens[i].Text == "h")) return true;
            break;
        }
        for (int i = index + 1; i < tokens.Count; i++)
        {
            if (tokens[i].Kind == Kind.Literal) continue;
            if (tokens[i].Kind == Kind.Second || (tokens[i].Kind == Kind.Elapsed && tokens[i].Text == "s")) return true;
            break;
        }
        return false;
    }
}
