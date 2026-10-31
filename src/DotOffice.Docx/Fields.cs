using System.Globalization;
using System.Text;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>A parsed field instruction (the text of w:instrText or w:fldSimple/@w:instr).</summary>
internal sealed record FieldInstruction(string Name, IReadOnlyList<string> Arguments, IReadOnlyDictionary<char, string?> Switches)
{
    /// <summary>Field we compute ourselves, or null when the cached result is shown as-is.</summary>
    public FieldKind? Kind => Name switch
    {
        "PAGE" => FieldKind.Page,
        "NUMPAGES" => FieldKind.NumPages,
        "SECTIONPAGES" => FieldKind.SectionPages,
        "PAGEREF" when Bookmark is not null => FieldKind.PageRef,
        _ => null,
    };

    /// <summary>The cached result is replaced by a <see cref="FieldRun"/> or a REF placeholder.</summary>
    public bool IsEvaluated => Kind is not null || (Name == "REF" && Bookmark is not null);

    public string? Bookmark => Name is "PAGEREF" or "REF" && Arguments.Count > 0 ? Arguments[0] : null;

    /// <summary>Link for HYPERLINK fields and for PAGEREF/REF with the \h switch.</summary>
    public Hyperlink? Link => Name switch
    {
        "HYPERLINK" when Arguments.Count > 0 || Switches.ContainsKey('l') =>
            new Hyperlink(Arguments.Count > 0 ? Arguments[0] : null, Switches.GetValueOrDefault('l')),
        "PAGEREF" or "REF" when Bookmark is not null && Switches.ContainsKey('h') => new Hyperlink(null, Bookmark),
        _ => null,
    };

    /// <summary>Date-time picture from the \@ switch.</summary>
    public string? Picture => Switches.GetValueOrDefault('@');

    /// <summary>Numeric format from the \* switch (MERGEFORMAT and other text formats are ignored).</summary>
    public PageNumberFormat? NumberFormat => Switches.GetValueOrDefault('*') switch
    {
        "Arabic" or "ARABIC" or "arabic" => PageNumberFormat.Decimal,
        "roman" => PageNumberFormat.LowerRoman,
        "ROMAN" or "Roman" => PageNumberFormat.UpperRoman,
        "alphabetic" => PageNumberFormat.LowerLetter,
        "ALPHABETIC" or "Alphabetic" => PageNumberFormat.UpperLetter,
        _ => null,
    };

    public Inline CreateInline(RunProperties props) => Kind is { } kind
        ? new FieldRun(kind, props) { Bookmark = Bookmark, NumberFormat = NumberFormat }
        : new RefPlaceholder(Bookmark!, props);

    public static FieldInstruction Parse(string instruction)
    {
        var tokens = Tokenize(instruction);
        var name = tokens.Count > 0 && !tokens[0].Switch ? tokens[0].Text.ToUpperInvariant() : "";
        var arguments = new List<string>();
        var switches = new Dictionary<char, string?>();
        for (int i = name.Length > 0 ? 1 : 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (!token.Switch)
            {
                arguments.Add(token.Text);
                continue;
            }
            // Switches with an argument: \@ "picture", \* format, \l "anchor", \o "target"…
            string? value = null;
            if (token.Text[0] is '@' or '*' or 'l' or 'o' or 't' or 'b' or 'f' or 'p' && i + 1 < tokens.Count && !tokens[i + 1].Switch)
                value = tokens[++i].Text;
            // \* MERGEFORMAT/CHARFORMAT only concern formatting on update; another \* switch may hold the number format.
            if (token.Text[0] == '*' && value is "MERGEFORMAT" or "CHARFORMAT") continue;
            switches.TryAdd(token.Text[0], value);
        }
        return new FieldInstruction(name, arguments, switches);
    }

    private static List<(string Text, bool Switch)> Tokenize(string s)
    {
        var tokens = new List<(string, bool)>();
        int i = 0;
        while (i < s.Length)
        {
            if (char.IsWhiteSpace(s[i])) { i++; continue; }
            if (s[i] == '"')
            {
                var sb = new StringBuilder();
                for (i++; i < s.Length && s[i] != '"'; i++)
                {
                    if (s[i] == '\\' && i + 1 < s.Length && s[i + 1] is '"' or '\\') i++;
                    sb.Append(s[i]);
                }
                i++;
                tokens.Add((sb.ToString(), false));
                continue;
            }
            if (s[i] == '\\' && i + 1 < s.Length)
            {
                tokens.Add((s[i + 1].ToString(), true));
                i += 2;
                continue;
            }
            int start = i;
            while (i < s.Length && !char.IsWhiteSpace(s[i]) && s[i] != '"') i++;
            tokens.Add((s[start..i], false));
        }
        return tokens;
    }

    /// <summary>Formats a date with a Word date-time picture (d, dd, ddd, dddd, M…MMMM, yy, yyyy, h, H, m, s, am/pm).</summary>
    public static string FormatDate(DateTime value, string picture)
    {
        var culture = CultureInfo.InvariantCulture;
        var sb = new StringBuilder();
        int i = 0;
        while (i < picture.Length)
        {
            char c = picture[i];
            if (c == '\'')
            {
                int close = picture.IndexOf('\'', i + 1);
                if (close < 0) close = picture.Length;
                sb.Append(picture, i + 1, close - i - 1);
                i = close + 1;
                continue;
            }
            if (string.Compare(picture, i, "am/pm", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                var text = value.Hour < 12 ? "am" : "pm";
                sb.Append(char.IsUpper(c) ? text.ToUpperInvariant() : text);
                i += 5;
                continue;
            }
            int n = 1;
            while (i + n < picture.Length && picture[i + n] == c) n++;
            switch (c)
            {
                case 'd' or 'D':
                    sb.Append(n switch
                    {
                        1 => value.Day.ToString(culture),
                        2 => value.Day.ToString("00", culture),
                        3 => value.ToString("ddd", culture),
                        _ => value.ToString("dddd", culture),
                    });
                    break;
                case 'M':
                    sb.Append(n switch
                    {
                        1 => value.Month.ToString(culture),
                        2 => value.Month.ToString("00", culture),
                        3 => value.ToString("MMM", culture),
                        _ => value.ToString("MMMM", culture),
                    });
                    break;
                case 'y' or 'Y':
                    sb.Append(n <= 2 ? (value.Year % 100).ToString("00", culture) : value.Year.ToString(culture));
                    break;
                case 'h':
                    int h12 = value.Hour % 12 == 0 ? 12 : value.Hour % 12;
                    sb.Append(n == 1 ? h12.ToString(culture) : h12.ToString("00", culture));
                    break;
                case 'H':
                    sb.Append(n == 1 ? value.Hour.ToString(culture) : value.Hour.ToString("00", culture));
                    break;
                case 'm':
                    sb.Append(n == 1 ? value.Minute.ToString(culture) : value.Minute.ToString("00", culture));
                    break;
                case 's' or 'S':
                    sb.Append(n == 1 ? value.Second.ToString(culture) : value.Second.ToString("00", culture));
                    break;
                default:
                    sb.Append(c, n);
                    break;
            }
            i += n;
        }
        return sb.ToString();
    }
}

/// <summary>A REF field; replaced by the bookmarked text once the whole document has been read.</summary>
internal sealed class RefPlaceholder(string bookmark, RunProperties properties) : Inline(properties)
{
    public string Bookmark { get; } = bookmark;
    public string? CachedResult { get; set; }
}
