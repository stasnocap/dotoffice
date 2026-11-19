namespace DotOffice.Docx;

/// <summary>
/// Maps characters of legacy symbol fonts (Symbol, Wingdings) to Unicode. Word stores such glyphs
/// as U+F0xx private-use code points (or plain bytes), which only render with the original
/// symbol-encoded font; mapping them lets any Unicode font draw bullets and arrows.
/// </summary>
internal static class SymbolFonts
{
    private static readonly Dictionary<int, char> Symbol = new()
    {
        [0xB7] = '•', [0xA7] = '♣', [0xA8] = '♦', [0xA9] = '♥', [0xAA] = '♠',
        [0xAB] = '↔', [0xAC] = '←', [0xAD] = '↑', [0xAE] = '→', [0xAF] = '↓',
        [0xB0] = '°', [0xB1] = '±', [0xB4] = '×', [0xB8] = '÷', [0xB9] = '≠',
        [0xBA] = '≡', [0xBB] = '≈', [0xBC] = '…', [0xD6] = '√', [0xD7] = '⋅',
        [0xE0] = '◊', [0x2D] = '−', [0x2A] = '∗', [0x6F] = 'ο',
        [0x61] = 'α', [0x62] = 'β', [0x63] = 'χ', [0x64] = 'δ', [0x65] = 'ε',
        [0x67] = 'γ', [0x6C] = 'λ', [0x6D] = 'μ', [0x70] = 'π', [0x73] = 'σ',
        [0x44] = 'Δ', [0x57] = 'Ω', [0x53] = 'Σ', [0x50] = 'Π',
    };

    private static readonly Dictionary<int, char> Wingdings = new()
    {
        [0x9F] = '•', [0xA7] = '▪', [0xA8] = '◻', [0x6C] = '●', [0x6D] = '❍',
        [0x6E] = '■', [0x6F] = '□', [0x70] = '◻', [0x71] = '❑', [0x72] = '❒',
        [0x73] = '⬧', [0x74] = '⧫', [0x75] = '◆', [0x76] = '❖', [0x77] = '⬥',
        [0xD8] = '➢', [0xE0] = '→', [0xE8] = '➔', [0xF0] = '⇨', [0xFC] = '✔',
        [0xFB] = '✘', [0xFD] = '☒', [0xFE] = '☑', [0xA1] = '○', [0xA4] = '◉',
        [0x4A] = '☺', [0x4B] = '☐', [0x4C] = '☹', [0x3F] = '✍', [0x43] = '☜',
        [0x46] = '☞', [0xAB] = '★', [0xB2] = '✧', [0xF6] = '⬈',
    };

    public static bool IsSymbolFont(string? font) =>
        font is not null && (font.Equals("Symbol", StringComparison.OrdinalIgnoreCase)
                             || font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase)
                             || font.Equals("Webdings", StringComparison.OrdinalIgnoreCase));

    /// <summary>Maps text written in a symbol font. Unknown symbols become <paramref name="unknown"/> (if given).</summary>
    public static string Map(string font, string text, char? unknown = null)
    {
        var table = font.Equals("Symbol", StringComparison.OrdinalIgnoreCase) ? Symbol : Wingdings;
        var chars = text.ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            int code = chars[i] is >= '' and <= '' ? chars[i] - 0xF000 : chars[i];
            if (code > 0xFF) continue;
            if (table.TryGetValue(code, out var mapped)) chars[i] = mapped;
            else if (code >= 0x20 && unknown is { } u) chars[i] = u;
            else if (chars[i] >= '') chars[i] = (char)code;
        }
        return new string(chars);
    }
}