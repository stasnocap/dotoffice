namespace DotOffice.Layout.Bidi;

/// <summary>Bidi_Class values (UAX #9, table 4). Order matches tools/generate-bidi-data.py.</summary>
internal enum BidiClass : byte { L, R, AL, EN, ES, ET, AN, CS, NSM, BN, B, S, WS, ON, LRE, LRO, RLE, RLO, PDF, LRI, RLI, FSI, PDI }

/// <summary>
/// The Unicode Bidirectional Algorithm (UAX #9): resolves embedding levels for one paragraph
/// (rules X1–X10, W1–W7, N0–N2, I1–I2 and L1) and reorders levels into visual order (L2).
/// Paragraph splitting (P1) is done by the caller; the paragraph level is given (P2/P3 are only
/// used for first-strong isolates).
/// </summary>
internal static class BidiAlgorithm
{
    private const int MaxDepth = 125;

    public static BidiClass GetClass(int codePoint)
    {
        var starts = BidiData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= codePoint) lo = mid; else hi = mid - 1;
        }
        return (BidiClass)BidiData.RangeValues[lo];
    }

    /// <summary>The Bidi_Mirroring_Glyph of <paramref name="codePoint"/>, or the code point itself.</summary>
    public static int Mirror(int codePoint)
    {
        var pairs = BidiData.Mirrors;
        int lo = 0, hi = pairs.Length / 2 - 1;
        while (lo <= hi)
        {
            int mid = (lo + hi) >> 1;
            int cp = pairs[mid * 2];
            if (cp == codePoint) return pairs[mid * 2 + 1];
            if (cp < codePoint) lo = mid + 1; else hi = mid - 1;
        }
        return codePoint;
    }

    /// <summary>Paired bracket of <paramref name="codePoint"/> (0 when none) and whether it opens.</summary>
    private static (int Pair, bool Opening) Bracket(int codePoint)
    {
        var brackets = BidiData.Brackets;
        int lo = 0, hi = brackets.Length / 3 - 1;
        while (lo <= hi)
        {
            int mid = (lo + hi) >> 1;
            int cp = brackets[mid * 3];
            if (cp == codePoint) return (brackets[mid * 3 + 1], brackets[mid * 3 + 2] == 1);
            if (cp < codePoint) lo = mid + 1; else hi = mid - 1;
        }
        return (0, false);
    }

    /// <summary>Direction of the first strong character (P2/P3): 1 for R/AL, 0 for L, -1 when there is none.</summary>
    public static int FirstStrong(ReadOnlySpan<BidiClass> classes)
    {
        int isolates = 0;
        foreach (var c in classes)
        {
            switch (c)
            {
                case BidiClass.LRI or BidiClass.RLI or BidiClass.FSI: isolates++; break;
                case BidiClass.PDI: if (isolates > 0) isolates--; break;
                case BidiClass.L when isolates == 0: return 0;
                case BidiClass.R or BidiClass.AL when isolates == 0: return 1;
                case BidiClass.B: return -1;
            }
        }
        return -1;
    }

    /// <summary>
    /// Resolves the embedding level of every character of a paragraph. <paramref name="codePoints"/> is
    /// used for bracket pairs (N0) and may contain 0 for characters inserted by the caller.
    /// </summary>
    public static byte[] ResolveLevels(ReadOnlySpan<BidiClass> classes, ReadOnlySpan<int> codePoints, byte paragraphLevel)
    {
        int n = classes.Length;
        var types = classes.ToArray();
        var levels = new byte[n];
        var matchingPdi = MatchIsolates(classes, out var matchingInitiator);
        ExplicitLevels(classes, types, levels, matchingPdi, paragraphLevel);

        foreach (var sequence in IsolatingRunSequences(classes, levels, matchingPdi, matchingInitiator, paragraphLevel))
            sequence.Resolve(types, levels, classes, codePoints);

        // Removed characters (X9) take the level of the preceding character so they reorder with it.
        for (int i = 0; i < n; i++)
            if (IsRemoved(classes[i])) levels[i] = i > 0 ? levels[i - 1] : paragraphLevel;

        // L1: separators, and whitespace before them or at the end of the paragraph, get the paragraph level.
        bool trailing = true;
        for (int i = n - 1; i >= 0; i--)
        {
            var c = classes[i];
            if (c is BidiClass.S or BidiClass.B)
            {
                levels[i] = paragraphLevel;
                trailing = true;
            }
            else if (trailing && (c is BidiClass.WS or BidiClass.FSI or BidiClass.LRI or BidiClass.RLI or BidiClass.PDI || IsRemoved(c)))
                levels[i] = paragraphLevel;
            else trailing = false;
        }
        return levels;
    }

    /// <summary>L2: indices of <paramref name="levels"/> in visual (left-to-right) order.</summary>
    public static int[] VisualOrder(ReadOnlySpan<byte> levels)
    {
        var order = new int[levels.Length];
        for (int i = 0; i < order.Length; i++) order[i] = i;
        if (levels.Length == 0) return order;
        byte max = 0, minOdd = byte.MaxValue;
        foreach (var level in levels)
        {
            max = Math.Max(max, level);
            if ((level & 1) == 1) minOdd = Math.Min(minOdd, level);
        }
        for (int level = max; level >= minOdd && level > 0; level--)
        {
            for (int i = 0; i < levels.Length; i++)
            {
                if (levels[order[i]] < level) continue;
                // order[] is permuted, so compare with the levels of the characters now at each position.
                int end = i;
                while (end + 1 < levels.Length && levels[order[end + 1]] >= level) end++;
                Array.Reverse(order, i, end - i + 1);
                i = end;
            }
        }
        return order;
    }

    private static bool IsRemoved(BidiClass c) =>
        c is BidiClass.RLE or BidiClass.LRE or BidiClass.RLO or BidiClass.LRO or BidiClass.PDF or BidiClass.BN;

    private static bool IsIsolateInitiator(BidiClass c) => c is BidiClass.LRI or BidiClass.RLI or BidiClass.FSI;

    /// <summary>BD9: the matching PDI of each isolate initiator (and the reverse), or -1.</summary>
    private static int[] MatchIsolates(ReadOnlySpan<BidiClass> classes, out int[] matchingInitiator)
    {
        var pdi = new int[classes.Length];
        matchingInitiator = new int[classes.Length];
        Array.Fill(pdi, -1);
        Array.Fill(matchingInitiator, -1);
        var open = new Stack<int>();
        for (int i = 0; i < classes.Length; i++)
        {
            if (IsIsolateInitiator(classes[i])) open.Push(i);
            else if (classes[i] == BidiClass.PDI && open.Count > 0)
            {
                int start = open.Pop();
                pdi[start] = i;
                matchingInitiator[i] = start;
            }
            else if (classes[i] == BidiClass.B) open.Clear();
        }
        return pdi;
    }

    private readonly record struct Entry(byte Level, BidiClass Override, bool Isolate);

    /// <summary>X1–X8.</summary>
    private static void ExplicitLevels(ReadOnlySpan<BidiClass> classes, BidiClass[] types, byte[] levels, int[] matchingPdi, byte paragraphLevel)
    {
        var stack = new Stack<Entry>();
        stack.Push(new Entry(paragraphLevel, BidiClass.ON, false));
        int overflowIsolates = 0, overflowEmbeddings = 0, validIsolates = 0;

        static byte NextOdd(byte level) => (byte)((level + 1) | 1);
        static byte NextEven(byte level) => (byte)((level + 2) & ~1);

        for (int i = 0; i < classes.Length; i++)
        {
            var top = stack.Peek();
            var c = classes[i];
            switch (c)
            {
                case BidiClass.RLE or BidiClass.LRE or BidiClass.RLO or BidiClass.LRO:
                {
                    byte level = c is BidiClass.RLE or BidiClass.RLO ? NextOdd(top.Level) : NextEven(top.Level);
                    if (level <= MaxDepth && overflowIsolates == 0 && overflowEmbeddings == 0)
                        stack.Push(new Entry(level, c switch { BidiClass.RLO => BidiClass.R, BidiClass.LRO => BidiClass.L, _ => BidiClass.ON }, false));
                    else if (overflowIsolates == 0) overflowEmbeddings++;
                    levels[i] = top.Level;
                    break;
                }
                case BidiClass.RLI or BidiClass.LRI or BidiClass.FSI:
                {
                    levels[i] = top.Level;
                    if (top.Override != BidiClass.ON) types[i] = top.Override;
                    bool rtl = c == BidiClass.RLI;
                    if (c == BidiClass.FSI)
                    {
                        int end = matchingPdi[i] < 0 ? classes.Length : matchingPdi[i];
                        rtl = FirstStrong(classes[(i + 1)..end]) == 1;
                    }
                    byte level = rtl ? NextOdd(top.Level) : NextEven(top.Level);
                    if (level <= MaxDepth && overflowIsolates == 0 && overflowEmbeddings == 0)
                    {
                        validIsolates++;
                        stack.Push(new Entry(level, BidiClass.ON, true));
                    }
                    else overflowIsolates++;
                    break;
                }
                case BidiClass.PDI:
                    if (overflowIsolates > 0) overflowIsolates--;
                    else if (validIsolates > 0)
                    {
                        overflowEmbeddings = 0;
                        while (!stack.Peek().Isolate) stack.Pop();
                        stack.Pop();
                        validIsolates--;
                    }
                    top = stack.Peek();
                    levels[i] = top.Level;
                    if (top.Override != BidiClass.ON) types[i] = top.Override;
                    break;
                case BidiClass.PDF:
                    if (overflowIsolates == 0)
                    {
                        if (overflowEmbeddings > 0) overflowEmbeddings--;
                        else if (!top.Isolate && stack.Count >= 2) stack.Pop();
                    }
                    levels[i] = top.Level;
                    break;
                case BidiClass.B:
                    levels[i] = paragraphLevel;
                    break;
                default:
                    levels[i] = top.Level;
                    if (top.Override != BidiClass.ON && c != BidiClass.BN) types[i] = top.Override;
                    break;
            }
        }
    }

    /// <summary>BD13/X10: level runs linked across isolates into isolating run sequences.</summary>
    private static List<Sequence> IsolatingRunSequences(ReadOnlySpan<BidiClass> classes, byte[] levels, int[] matchingPdi,
        int[] matchingInitiator, byte paragraphLevel)
    {
        int n = classes.Length;
        var runs = new List<List<int>>();
        var runOf = new int[n];
        List<int>? current = null;
        int lastLevel = -1;
        for (int i = 0; i < n; i++)
        {
            if (IsRemoved(classes[i])) { runOf[i] = -1; continue; }
            if (current is null || levels[i] != lastLevel)
            {
                current = [];
                runs.Add(current);
                lastLevel = levels[i];
            }
            current.Add(i);
            runOf[i] = runs.Count - 1;
        }

        var sequences = new List<Sequence>();
        foreach (var run in runs)
        {
            int first = run[0];
            if (classes[first] == BidiClass.PDI && matchingInitiator[first] >= 0) continue; // continues an earlier sequence
            var indexes = new List<int>(run);
            while (true)
            {
                int last = indexes[^1];
                if (!IsIsolateInitiator(classes[last]) || matchingPdi[last] < 0) break;
                indexes.AddRange(runs[runOf[matchingPdi[last]]]);
            }

            byte level = levels[indexes[0]];
            int before = indexes[0] - 1;
            while (before >= 0 && IsRemoved(classes[before])) before--;
            byte prevLevel = before >= 0 ? levels[before] : paragraphLevel;
            int lastIndex = indexes[^1];
            byte nextLevel = paragraphLevel;
            if (!IsIsolateInitiator(classes[lastIndex])) // an initiator ending a sequence has no matching PDI
            {
                int after = lastIndex + 1;
                while (after < n && IsRemoved(classes[after])) after++;
                if (after < n) nextLevel = levels[after];
            }
            sequences.Add(new Sequence(indexes.ToArray(), level,
                (Math.Max(prevLevel, level) & 1) == 1 ? BidiClass.R : BidiClass.L,
                (Math.Max(nextLevel, level) & 1) == 1 ? BidiClass.R : BidiClass.L));
        }
        return sequences;
    }

    private sealed class Sequence(int[] indexes, byte level, BidiClass sos, BidiClass eos)
    {
        private BidiClass[] _t = [];

        private BidiClass Embedding => (level & 1) == 1 ? BidiClass.R : BidiClass.L;

        public void Resolve(BidiClass[] types, byte[] levels, ReadOnlySpan<BidiClass> original, ReadOnlySpan<int> codePoints)
        {
            _t = indexes.Select(i => types[i]).ToArray();
            ResolveWeak();
            ResolveBrackets(original, codePoints);
            ResolveNeutral();
            for (int k = 0; k < indexes.Length; k++)
            {
                int i = indexes[k];
                var t = _t[k];
                types[i] = t;
                // I1/I2
                if ((level & 1) == 0)
                    levels[i] = (byte)(level + (t == BidiClass.R ? 1 : t is BidiClass.AN or BidiClass.EN ? 2 : 0));
                else
                    levels[i] = (byte)(level + (t is BidiClass.L or BidiClass.EN or BidiClass.AN ? 1 : 0));
            }
        }

        private void ResolveWeak()
        {
            var t = _t;
            int n = t.Length;
            // W1
            for (int k = 0; k < n; k++)
            {
                if (t[k] != BidiClass.NSM) continue;
                var prev = k == 0 ? sos : t[k - 1];
                t[k] = IsIsolateInitiator(prev) || prev == BidiClass.PDI ? BidiClass.ON : prev;
            }
            // W2, W3
            var lastStrong = sos;
            for (int k = 0; k < n; k++)
            {
                if (t[k] is BidiClass.L or BidiClass.R or BidiClass.AL) lastStrong = t[k];
                else if (t[k] == BidiClass.EN && lastStrong == BidiClass.AL) t[k] = BidiClass.AN;
            }
            for (int k = 0; k < n; k++)
                if (t[k] == BidiClass.AL) t[k] = BidiClass.R;
            // W4
            for (int k = 1; k + 1 < n; k++)
            {
                if (t[k] == BidiClass.ES && t[k - 1] == BidiClass.EN && t[k + 1] == BidiClass.EN) t[k] = BidiClass.EN;
                else if (t[k] == BidiClass.CS && t[k - 1] == t[k + 1] && t[k - 1] is BidiClass.EN or BidiClass.AN) t[k] = t[k - 1];
            }
            // W5
            for (int k = 0; k < n; k++)
            {
                if (t[k] != BidiClass.ET) continue;
                int end = k;
                while (end < n && t[end] == BidiClass.ET) end++;
                bool adjacent = (k > 0 && t[k - 1] == BidiClass.EN) || (end < n && t[end] == BidiClass.EN);
                if (adjacent) for (int m = k; m < end; m++) t[m] = BidiClass.EN;
                k = end - 1;
            }
            // W6
            for (int k = 0; k < n; k++)
                if (t[k] is BidiClass.ES or BidiClass.ET or BidiClass.CS) t[k] = BidiClass.ON;
            // W7
            lastStrong = sos;
            for (int k = 0; k < n; k++)
            {
                if (t[k] is BidiClass.L or BidiClass.R) lastStrong = t[k];
                else if (t[k] == BidiClass.EN && lastStrong == BidiClass.L) t[k] = BidiClass.L;
            }
        }

        /// <summary>Strong direction for N0/N1: EN and AN count as R.</summary>
        private static BidiClass? Strong(BidiClass t) => t switch
        {
            BidiClass.L => BidiClass.L,
            BidiClass.R or BidiClass.AL or BidiClass.EN or BidiClass.AN => BidiClass.R,
            _ => null,
        };

        private void ResolveBrackets(ReadOnlySpan<BidiClass> original, ReadOnlySpan<int> codePoints)
        {
            // BD16: find bracket pairs.
            var pairs = new List<(int Open, int Close)>();
            var stack = new List<(int Pair, int Position)>();
            for (int k = 0; k < indexes.Length; k++)
            {
                if (_t[k] != BidiClass.ON) continue;
                int cp = Canonical(codePoints[indexes[k]]);
                var (pair, opening) = Bracket(cp);
                if (pair == 0) continue;
                if (opening)
                {
                    if (stack.Count == 63) break;
                    stack.Add((Canonical(pair), k));
                }
                else
                {
                    for (int s = stack.Count - 1; s >= 0; s--)
                    {
                        if (stack[s].Pair != cp) continue;
                        pairs.Add((stack[s].Position, k));
                        stack.RemoveRange(s, stack.Count - s);
                        break;
                    }
                }
            }
            pairs.Sort((a, b) => a.Open.CompareTo(b.Open));

            var embedding = Embedding;
            foreach (var (open, close) in pairs)
            {
                bool sameInside = false, oppositeInside = false;
                for (int k = open + 1; k < close; k++)
                {
                    if (Strong(_t[k]) is not { } dir) continue;
                    if (dir == embedding) { sameInside = true; break; }
                    oppositeInside = true;
                }
                BidiClass? resolved = null;
                if (sameInside) resolved = embedding;
                else if (oppositeInside)
                {
                    var context = sos;
                    for (int k = open - 1; k >= 0; k--)
                        if (Strong(_t[k]) is { } dir) { context = dir; break; }
                    resolved = context != embedding ? context : embedding;
                }
                if (resolved is not { } r) continue;
                _t[open] = _t[close] = r;
                foreach (int bracket in new[] { open, close })
                    for (int k = bracket + 1; k < indexes.Length && original[indexes[k]] == BidiClass.NSM; k++)
                        _t[k] = r;
            }
        }

        private static int Canonical(int cp) => cp switch { 0x2329 => 0x3008, 0x232A => 0x3009, _ => cp };

        private void ResolveNeutral()
        {
            var t = _t;
            int n = t.Length;
            for (int k = 0; k < n; k++)
            {
                if (!IsNeutral(t[k])) continue;
                int end = k;
                while (end < n && IsNeutral(t[end])) end++;
                var before = k == 0 ? sos : Strong(t[k - 1]) ?? Embedding;
                var after = end == n ? eos : Strong(t[end]) ?? Embedding;
                var resolved = before == after ? before : Embedding; // N1, N2
                for (int m = k; m < end; m++) t[m] = resolved;
                k = end - 1;
            }
        }

        private static bool IsNeutral(BidiClass c) =>
            c is BidiClass.B or BidiClass.S or BidiClass.WS or BidiClass.ON or BidiClass.BN
                or BidiClass.LRI or BidiClass.RLI or BidiClass.FSI or BidiClass.PDI;
    }
}
