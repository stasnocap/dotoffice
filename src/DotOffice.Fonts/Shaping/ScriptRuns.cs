namespace DotOffice.Fonts.Shaping;

/// <summary>
/// Splits text into runs of one script before shaping, as LibreOffice does (vcl's ScriptRun, from ICU's
/// scrptrun): Common and Inherited characters join the run around them, and a closing bracket takes the
/// script of its opening bracket.
/// </summary>
internal static class ScriptRuns
{
    /// <summary>The runs of <paramref name="text"/> as (UTF-16 start, length, OpenType script tag).</summary>
    public static List<(int Start, int Length, string Tag)> Split(string text)
    {
        var runs = new List<(int, int, string)>();
        var stack = new List<(int PairIndex, byte Script)>();
        int sp = -1, end = 0;
        while (end < text.Length)
        {
            int startSp = sp, start = end;
            byte script = ScriptData.Common;
            for (; end < text.Length; end++)
            {
                int ch = text[end];
                if (char.IsHighSurrogate(text[end]) && end < text.Length - 1 && char.IsLowSurrogate(text[end + 1]))
                    ch = char.ConvertToUtf32(text[end], text[++end]);

                byte sc = ScriptOf(ch);
                int pairIndex = PairIndex(ch);
                if (pairIndex >= 0)
                {
                    if ((pairIndex & 1) == 0)
                    {
                        sp++;
                        if (sp == stack.Count) stack.Add(default);
                        stack[sp] = (pairIndex, script);
                    }
                    else if (sp >= 0)
                    {
                        while (sp >= 0 && stack[sp].PairIndex != (pairIndex & ~1)) sp--;
                        if (sp < startSp) startSp = sp;
                        if (sp >= 0) sc = stack[sp].Script;
                    }
                }

                if (script <= ScriptData.Inherited || sc <= ScriptData.Inherited || script == sc)
                {
                    if (script <= ScriptData.Inherited && sc > ScriptData.Inherited)
                    {
                        script = sc;
                        // open brackets pushed before the script was known take it now
                        while (startSp < sp) { startSp++; stack[startSp] = stack[startSp] with { Script = script }; }
                    }
                    if (pairIndex >= 0 && (pairIndex & 1) != 0 && sp >= 0)
                    {
                        sp--;
                        if (startSp >= 0) startSp--;
                    }
                }
                else
                {
                    if (ch >= 0x10000) end--; // end the run before the surrogate pair
                    break;
                }
            }
            runs.Add((start, end - start, ScriptData.Tags[script]));
        }
        return runs;
    }

    /// <summary>The script index of a character, with LibreOffice's adjustments (see ScriptData).</summary>
    public static byte ScriptOf(int cp)
    {
        var starts = ScriptData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        return ScriptData.RangeScripts[lo];
    }

    /// <summary>Script tag of a character (DFLT for Common, Inherited and Unknown).</summary>
    public static string TagOf(int cp) => ScriptData.Tags[ScriptOf(cp)];

    /// <summary>Paired punctuation: even indices open, the following odd index closes.</summary>
    private static int PairIndex(int ch) => ch switch
    {
        '(' => 0, ')' => 1, '<' => 2, '>' => 3, '[' => 4, ']' => 5, '{' => 6, '}' => 7, '«' => 8, '»' => 9,
        '‘' => 10, '’' => 11, '“' => 12, '”' => 13, '‹' => 14, '›' => 15,
        >= 0x3008 and <= 0x3011 => 16 + ch - 0x3008,
        >= 0x3014 and <= 0x301B => 26 + ch - 0x3014,
        _ => -1,
    };
}
