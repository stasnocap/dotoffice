namespace DotOffice.Fonts.Shaping;

/// <summary>
/// Regular expressions over character categories (one byte each), for the syllable scanners of the Indic and
/// Universal Shaping Engine shapers. <see cref="Pattern.Longest"/> gives the longest match, like a Ragel scanner.
/// </summary>
internal static class SyllablePatterns
{
    public abstract class Pattern
    {
        /// <summary>Every position where a match starting at <paramref name="start"/> can end.</summary>
        public abstract IEnumerable<int> Ends(byte[] input, int start);

        /// <summary>End of the longest match at <paramref name="start"/>, or <paramref name="start"/> when there is none.</summary>
        public int Longest(byte[] input, int start) => Ends(input, start).DefaultIfEmpty(start).Max();
    }

    private sealed class Symbol(params byte[] categories) : Pattern
    {
        public override IEnumerable<int> Ends(byte[] input, int start)
        {
            if (start < input.Length && Array.IndexOf(categories, input[start]) >= 0) yield return start + 1;
        }
    }

    private sealed class Sequence(params Pattern[] parts) : Pattern
    {
        public override IEnumerable<int> Ends(byte[] input, int start)
        {
            IEnumerable<int> ends = [start];
            foreach (var part in parts) ends = ends.SelectMany(e => part.Ends(input, e)).Distinct().ToList();
            return ends;
        }
    }

    private sealed class Choice(params Pattern[] options) : Pattern
    {
        public override IEnumerable<int> Ends(byte[] input, int start) =>
            options.SelectMany(o => o.Ends(input, start)).Distinct();
    }

    private sealed class Repeat(Pattern item, int min) : Pattern
    {
        public override IEnumerable<int> Ends(byte[] input, int start)
        {
            var reached = new HashSet<int>();
            var frontier = new List<int> { start };
            for (int count = 0; frontier.Count > 0; count++)
            {
                if (count >= min) foreach (var e in frontier) reached.Add(e);
                frontier = frontier.SelectMany(e => item.Ends(input, e)).Where(e => count < min || !reached.Contains(e)).Distinct().ToList();
                if (count > input.Length) break;
            }
            return reached;
        }
    }

    public static Pattern S(params byte[] c) => new Symbol(c);
    public static Pattern Seq(params Pattern[] p) => new Sequence(p);
    public static Pattern Or(params Pattern[] p) => new Choice(p);
    public static Pattern Opt(Pattern p) => Or(Seq(), p);
    public static Pattern Star(Pattern p) => new Repeat(p, 0);
}
