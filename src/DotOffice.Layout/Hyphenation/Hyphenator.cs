using System.Reflection;

namespace DotOffice.Layout.Hyphenation;

/// <summary>
/// Liang's hyphenation algorithm over the pattern dictionaries LibreOffice ships (Hyphen's .dic format),
/// embedded per language so that more languages can be added by dropping in a file.
/// </summary>
internal sealed class Hyphenator
{
    private static readonly Dictionary<string, Hyphenator?> Cache = new(StringComparer.OrdinalIgnoreCase);

    private readonly Dictionary<string, byte[]> patterns = [];
    private readonly int maxPattern;

    /// <summary>Letters that must stay before the first and after the last hyphen.</summary>
    public int LeftMin { get; private set; } = 2;
    public int RightMin { get; private set; } = 2;

    private Hyphenator(TextReader reader)
    {
        reader.ReadLine(); // character set (always UTF-8 in the embedded files)
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            line = line.Trim();
            if (line.Length == 0 || line[0] == '%') continue;
            if (line.StartsWith("LEFTHYPHENMIN ")) { LeftMin = int.Parse(line[14..]); continue; }
            if (line.StartsWith("RIGHTHYPHENMIN ")) { RightMin = int.Parse(line[15..]); continue; }
            if (line.StartsWith("COMPOUND") || line.StartsWith("NEXTLEVEL")) continue;

            var letters = new List<char>();
            var digits = new List<byte>();
            byte pending = 0;
            foreach (char c in line)
            {
                if (c is >= '0' and <= '9') pending = (byte)(c - '0');
                else { digits.Add(pending); pending = 0; letters.Add(c); }
            }
            digits.Add(pending);
            patterns[new string([.. letters])] = [.. digits];
            maxPattern = Math.Max(maxPattern, letters.Count);
        }
    }

    /// <summary>The hyphenator for a language tag such as "en-US", or null when there are no patterns for it.</summary>
    public static Hyphenator? ForLanguage(string language)
    {
        lock (Cache)
        {
            if (Cache.TryGetValue(language, out var cached)) return cached;
            var assembly = typeof(Hyphenator).Assembly;
            using var stream = assembly.GetManifestResourceStream($"DotOffice.Hyphenation.hyph_{language.Replace('-', '_')}.dic");
            return Cache[language] = stream is null ? null : new Hyphenator(new StreamReader(stream));
        }
    }

    /// <summary>
    /// Offsets in <paramref name="word"/> (letters only) where a hyphen may be inserted: a hyphen goes before the
    /// character at each returned offset.
    /// </summary>
    public List<int> Find(string word)
    {
        var result = new List<int>();
        int n = word.Length;
        if (n < LeftMin + RightMin) return result;
        var padded = "." + word.ToLowerInvariant() + ".";
        var points = new byte[padded.Length + 1];
        for (int start = 0; start < padded.Length; start++)
            for (int len = 1; len <= maxPattern && start + len <= padded.Length; len++)
            {
                if (!patterns.TryGetValue(padded.Substring(start, len), out var digits)) continue;
                for (int k = 0; k < digits.Length; k++)
                    points[start + k] = Math.Max(points[start + k], digits[k]);
            }
        // points[j] sits before padded[j]; padded[1] is word[0].
        for (int offset = LeftMin; offset <= n - RightMin; offset++)
            if ((points[offset + 1] & 1) == 1) result.Add(offset);
        return result;
    }
}
