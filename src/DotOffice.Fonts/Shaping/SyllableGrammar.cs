using static DotOffice.Fonts.Shaping.IndicCategory;
using static DotOffice.Fonts.Shaping.SyllablePatterns;

namespace DotOffice.Fonts.Shaping;

/// <summary>
/// The Indic syllable grammar of HarfBuzz's indic_syllable_machine (hb-ot-shaper-indic-machine.rl), as
/// expressions over character categories. <see cref="Pattern.Longest"/> gives the longest match, like the
/// Ragel scanner.
/// </summary>
internal static class SyllableGrammar
{
    private static Pattern S(params IndicCategory[] c) => SyllablePatterns.S(c.Select(x => (byte)x).ToArray());
    private static Pattern Seq(params Pattern[] p) => SyllablePatterns.Seq(p);
    private static Pattern Or(params Pattern[] p) => SyllablePatterns.Or(p);
    private static Pattern Opt(Pattern p) => SyllablePatterns.Opt(p);
    private static Pattern Star(Pattern p) => SyllablePatterns.Star(p);

    private static readonly Pattern C = S(IndicCategory.C, Ra);
    private static readonly Pattern N1 = Seq(Opt(Seq(Opt(S(ZWNJ)), S(RS))), Opt(Seq(S(N), Opt(S(N)))));
    private static readonly Pattern Z = S(ZWJ, ZWNJ);
    private static readonly Pattern Reph = Or(Seq(S(Ra), S(H)), S(Repha));
    private static readonly Pattern Sm = S(SM, SMPst);
    private static readonly Pattern Cn = Seq(C, Opt(S(ZWJ)), Opt(N1));
    private static readonly Pattern SymbolPart = Seq(S(IndicCategory.Symbol), Opt(S(N)));
    private static readonly Pattern MatraGroup = Seq(Star(Z), Or(S(M), Seq(Opt(Sm), S(MPst))), Opt(S(N)), Opt(S(H)));
    private static readonly Pattern SyllableTail = Seq(Opt(Seq(Opt(Z), Sm, Opt(Sm), Opt(S(ZWNJ)))), Star(S(A)));
    private static readonly Pattern HalantGroup = Seq(Opt(Z), S(H), Opt(Seq(S(ZWJ), Opt(S(N)))));
    private static readonly Pattern FinalHalantGroup = Or(HalantGroup, Seq(S(H), S(ZWNJ)));
    private static readonly Pattern MedialGroup = Opt(S(CM));
    private static readonly Pattern HalantOrMatraGroup = Or(FinalHalantGroup, Star(MatraGroup));
    private static readonly Pattern ComplexSyllableTail = Seq(Star(Seq(HalantGroup, Cn)), MedialGroup, HalantOrMatraGroup, SyllableTail);

    /// <summary>The scanner's rules in priority order (an earlier rule wins a tie).</summary>
    internal static readonly (Pattern Rule, int Type)[] RuleList =
    [
        (Seq(Opt(S(Repha, CS)), Cn, ComplexSyllableTail), 0), // consonant syllable
        (Seq(Opt(Reph), S(V), Opt(N1), Or(S(ZWJ), ComplexSyllableTail)), 1), // vowel syllable
        (Seq(Or(Seq(Opt(S(Repha, CS)), S(Placeholder)), Seq(Opt(Reph), S(DottedCircle))), Opt(N1), ComplexSyllableTail), 2), // standalone
        (Seq(SymbolPart, SyllableTail), 3), // symbol cluster
        (S(SMPst), 5), // non-Indic
        (Seq(Opt(Reph), Opt(N1), ComplexSyllableTail), 4), // broken cluster
    ];
}
