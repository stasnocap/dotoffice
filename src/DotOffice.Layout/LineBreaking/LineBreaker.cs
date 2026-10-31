namespace DotOffice.Layout.LineBreaking;

/// <summary>Unicode Line_Break property values (UAX #14). Order must match tools/generate-linebreak-data.py.</summary>
internal enum LineBreakClass : byte
{
    XX, BK, CR, LF, CM, NL, SG, WJ, ZW, GL, SP, ZWJ, B2, BA, BB, HY, CB, CL, CP, EX, IN, NS, OP, QU, IS, NU, PO, PR, SY,
    AI, AL, CJ, EB, EM, H2, H3, HL, ID, JL, JV, JT, RI, SA, AK, AP, AS, VF, VI,
}

public enum LineBreakOpportunity : byte { None, Allowed, Mandatory }

/// <summary>
/// Line break opportunities per UAX #14 (Unicode Line Breaking Algorithm), implementing rules
/// LB1–LB31 with these simplifications: SA (Thai, Lao…) is treated as AL (no dictionary breaking);
/// LB15a–d, LB21a, LB25 (full regex form) and LB28a (Brahmic aksara) use their simple pair forms;
/// LB30 ignores East Asian Width. Line_Break data is generated from the Unicode Character Database.
/// </summary>
public static class LineBreaker
{
    /// <summary>
    /// Fills <paramref name="breaks"/> (same length as <paramref name="text"/>): entry i describes the
    /// opportunity to break <em>before</em> <c>text[i]</c>. Entry 0 is always None; the end of text is
    /// an implicit mandatory break.
    /// </summary>
    public static void FindBreaks(ReadOnlySpan<char> text, Span<LineBreakOpportunity> breaks)
    {
        if (breaks.Length < text.Length) throw new ArgumentException("Output span is too short.", nameof(breaks));
        breaks[..text.Length].Clear();
        if (text.IsEmpty) return;

        int i = 0;
        var first = Resolve(Next(text, ref i));
        var prevRaw = first;
        var prev = first is LineBreakClass.CM or LineBreakClass.ZWJ ? LineBreakClass.AL : first;
        var beforeSpaces = prev;
        int riCount = prev == LineBreakClass.RI ? 1 : 0;

        while (i < text.Length)
        {
            int position = i;
            var raw = Resolve(Next(text, ref i));
            var cur = raw;
            LineBreakOpportunity decision;

            if (prev is LineBreakClass.BK or LineBreakClass.LF or LineBreakClass.NL
                || (prev == LineBreakClass.CR && cur != LineBreakClass.LF))
                decision = LineBreakOpportunity.Mandatory;                                   // LB4, LB5
            else if (cur is LineBreakClass.BK or LineBreakClass.CR or LineBreakClass.LF or LineBreakClass.NL
                     || cur is LineBreakClass.SP or LineBreakClass.ZW)
                decision = LineBreakOpportunity.None;                                        // LB5, LB6, LB7
            else if (beforeSpaces == LineBreakClass.ZW && prev is LineBreakClass.ZW or LineBreakClass.SP)
                decision = LineBreakOpportunity.Allowed;                                     // LB8
            else if (cur is LineBreakClass.CM or LineBreakClass.ZWJ
                     && prev is not (LineBreakClass.BK or LineBreakClass.CR or LineBreakClass.LF or LineBreakClass.NL
                         or LineBreakClass.SP or LineBreakClass.ZW))
            {
                // LB9: combining marks (also after a ZWJ) take the class of their base; nothing else changes.
                breaks[position] = LineBreakOpportunity.None;
                prevRaw = raw;
                continue;
            }
            else if (prevRaw == LineBreakClass.ZWJ)
                decision = LineBreakOpportunity.None;                                        // LB8a
            else
            {
                if (cur is LineBreakClass.CM or LineBreakClass.ZWJ) cur = LineBreakClass.AL; // LB10
                decision = Pair(prev, cur, beforeSpaces, riCount);
            }

            breaks[position] = decision;
            riCount = cur == LineBreakClass.RI ? riCount + 1 : 0;
            prevRaw = raw;
            prev = cur;
            if (cur != LineBreakClass.SP) beforeSpaces = cur;
        }
    }

    /// <summary>Convenience wrapper returning the indices where a break may occur (excluding 0).</summary>
    public static List<(int Index, bool Mandatory)> GetBreaks(string text)
    {
        var breaks = new LineBreakOpportunity[text.Length];
        FindBreaks(text, breaks);
        var result = new List<(int, bool)>();
        for (int i = 1; i < breaks.Length; i++)
            if (breaks[i] != LineBreakOpportunity.None) result.Add((i, breaks[i] == LineBreakOpportunity.Mandatory));
        return result;
    }

    internal static LineBreakClass GetClass(int codePoint)
    {
        var starts = LineBreakData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= codePoint) lo = mid; else hi = mid - 1;
        }
        return (LineBreakClass)LineBreakData.RangeClasses[lo];
    }

    private static int Next(ReadOnlySpan<char> text, ref int i)
    {
        char c = text[i++];
        if (char.IsHighSurrogate(c) && i < text.Length && char.IsLowSurrogate(text[i]))
            return char.ConvertToUtf32(c, text[i++]);
        return c;
    }

    /// <summary>LB1 class resolution.</summary>
    private static LineBreakClass Resolve(int codePoint) => GetClass(codePoint) switch
    {
        LineBreakClass.AI or LineBreakClass.SG or LineBreakClass.XX or LineBreakClass.SA
            or LineBreakClass.AK or LineBreakClass.AP or LineBreakClass.AS
            or LineBreakClass.VF or LineBreakClass.VI => LineBreakClass.AL,
        LineBreakClass.CJ => LineBreakClass.NS,
        var c => c,
    };

    private static LineBreakOpportunity Pair(LineBreakClass prev, LineBreakClass cur, LineBreakClass beforeSpaces, int riCount)
    {
        const LineBreakOpportunity no = LineBreakOpportunity.None, yes = LineBreakOpportunity.Allowed;
        if (cur == LineBreakClass.WJ || prev == LineBreakClass.WJ) return no;                            // LB11
        if (prev == LineBreakClass.GL) return no;                                                        // LB12
        if (cur == LineBreakClass.GL && prev is not (LineBreakClass.SP or LineBreakClass.BA or LineBreakClass.HY))
            return no;                                                                                   // LB12a
        if (cur is LineBreakClass.CL or LineBreakClass.CP or LineBreakClass.EX or LineBreakClass.IS or LineBreakClass.SY)
            return no;                                                                                   // LB13
        if (beforeSpaces == LineBreakClass.OP) return no;                                                // LB14
        if (beforeSpaces == LineBreakClass.QU && cur == LineBreakClass.OP) return no;                    // LB15
        if (beforeSpaces is LineBreakClass.CL or LineBreakClass.CP && cur == LineBreakClass.NS) return no; // LB16
        if (beforeSpaces == LineBreakClass.B2 && cur == LineBreakClass.B2) return no;                    // LB17
        if (prev == LineBreakClass.SP) return yes;                                                       // LB18
        if (cur == LineBreakClass.QU || prev == LineBreakClass.QU) return no;                            // LB19
        if (cur == LineBreakClass.CB || prev == LineBreakClass.CB) return yes;                           // LB20
        if (cur is LineBreakClass.BA or LineBreakClass.HY or LineBreakClass.NS || prev == LineBreakClass.BB)
            return no;                                                                                   // LB21
        if (prev == LineBreakClass.SY && cur == LineBreakClass.HL) return no;                            // LB21b
        if (cur == LineBreakClass.IN) return no;                                                         // LB22
        if (IsAlpha(prev) && cur == LineBreakClass.NU || prev == LineBreakClass.NU && IsAlpha(cur)) return no; // LB23
        if (prev == LineBreakClass.PR && cur is LineBreakClass.ID or LineBreakClass.EB or LineBreakClass.EM
            || prev is LineBreakClass.ID or LineBreakClass.EB or LineBreakClass.EM && cur == LineBreakClass.PO)
            return no;                                                                                   // LB23a
        if (prev is LineBreakClass.PR or LineBreakClass.PO && IsAlpha(cur)
            || IsAlpha(prev) && cur is LineBreakClass.PR or LineBreakClass.PO)
            return no;                                                                                   // LB24
        if (prev is LineBreakClass.CL or LineBreakClass.CP or LineBreakClass.NU && cur is LineBreakClass.PO or LineBreakClass.PR
            || prev is LineBreakClass.PO or LineBreakClass.PR && cur is LineBreakClass.OP or LineBreakClass.NU
            || prev is LineBreakClass.HY or LineBreakClass.IS or LineBreakClass.NU or LineBreakClass.SY && cur == LineBreakClass.NU)
            return no;                                                                                   // LB25
        if (prev == LineBreakClass.JL && cur is LineBreakClass.JL or LineBreakClass.JV or LineBreakClass.H2 or LineBreakClass.H3
            || prev is LineBreakClass.JV or LineBreakClass.H2 && cur is LineBreakClass.JV or LineBreakClass.JT
            || prev is LineBreakClass.JT or LineBreakClass.H3 && cur == LineBreakClass.JT)
            return no;                                                                                   // LB26
        if (IsKorean(prev) && cur == LineBreakClass.PO || prev == LineBreakClass.PR && IsKorean(cur)) return no; // LB27
        if (IsAlpha(prev) && IsAlpha(cur)) return no;                                                    // LB28
        if (prev == LineBreakClass.IS && IsAlpha(cur)) return no;                                        // LB29
        if ((IsAlpha(prev) || prev == LineBreakClass.NU) && cur == LineBreakClass.OP
            || prev == LineBreakClass.CP && (IsAlpha(cur) || cur == LineBreakClass.NU))
            return no;                                                                                   // LB30
        if (prev == LineBreakClass.RI && cur == LineBreakClass.RI && riCount % 2 == 1) return no;        // LB30a
        if (prev == LineBreakClass.EB && cur == LineBreakClass.EM) return no;                            // LB30b
        return yes;                                                                                      // LB31
    }

    private static bool IsAlpha(LineBreakClass c) => c is LineBreakClass.AL or LineBreakClass.HL;

    private static bool IsKorean(LineBreakClass c) =>
        c is LineBreakClass.JL or LineBreakClass.JV or LineBreakClass.JT or LineBreakClass.H2 or LineBreakClass.H3;
}