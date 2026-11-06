using DotOffice.Fonts.Shaping;

namespace DotOffice.Layout;

/// <summary>
/// Text running down (<c>tbRl</c>) is laid out horizontally in a frame that is then turned a quarter clockwise. East Asian characters stay
/// upright in such text while Latin letters are turned with the line, so each of those characters is drawn in a group that turns it
/// back about the center of its em square.
/// </summary>
internal static class UprightText
{
    /// <summary>Where the center of an ideographic em square lies above the baseline, in ems (ascent 0.88, descent 0.12).</summary>
    private const double CenterAboveBaseline = 0.38;

    internal static bool IsUpright(int codePoint) =>
        codePoint is >= 0x2E80 and <= 0xA4CF or >= 0xAC00 and <= 0xD7AF or >= 0xF900 and <= 0xFAFF or >= 0xFE30 and <= 0xFE4F
            or >= 0xFF00 and <= 0xFF60 or >= 0xFFE0 and <= 0xFFE6 or >= 0x20000 and <= 0x3FFFF;

    /// <summary>
    /// Returns <paramref name="items"/> (in the turned frame) with every upright character in its own counter-turned group. Items inside a
    /// group that itself turns or mirrors (a picture, a nested vertical cell) are left alone.
    /// </summary>
    internal static List<LayoutItem> Apply(IEnumerable<LayoutItem> items)
    {
        var result = new List<LayoutItem>();
        var turning = new Stack<bool>();
        int turned = 0;
        foreach (var item in items)
        {
            switch (item)
            {
                case PushItem push:
                    bool t = push.Rotation != 0 || push.FlipH || push.FlipV;
                    turning.Push(t);
                    if (t) turned++;
                    result.Add(push);
                    break;
                case PopItem:
                    if (turning.Count > 0 && turning.Pop()) turned--;
                    result.Add(item);
                    break;
                case TextItem { RightToLeft: false, Field: null } text when turned == 0 && Split(text, result):
                    break;
                default:
                    result.Add(item);
                    break;
            }
        }
        return result;
    }

    /// <summary>Adds the pieces of <paramref name="text"/> to <paramref name="output"/>; false (adding nothing) when it holds no upright character.</summary>
    private static bool Split(TextItem text, List<LayoutItem> output)
    {
        var pieces = new List<(string Text, IReadOnlyList<ShapedGlyph>? Glyphs, double X, double Width, bool Upright)>();
        double scale = text.FontSize / text.Font.UnitsPerEm;
        bool any = false;
        if (text.Glyphs is { } glyphs)
        {
            double x = text.X;
            var run = new List<ShapedGlyph>();
            double runX = x, runWidth = 0;
            void Flush()
            {
                if (run.Count > 0) pieces.Add((string.Concat(run.Select(g => g.Text)), run.ToList(), runX, runWidth, false));
                run.Clear();
                runWidth = 0;
            }
            foreach (var g in glyphs)
            {
                double advance = g.Advance * scale + (g.Advance != 0 ? text.CharacterSpacing : 0);
                bool upright = g.Text.Length > 0 && IsUpright(char.ConvertToUtf32(g.Text, 0));
                if (upright)
                {
                    Flush();
                    pieces.Add((g.Text, new[] { g }, x, advance, true));
                    any = true;
                    runX = x + advance;
                }
                else
                {
                    if (run.Count == 0) runX = x;
                    run.Add(g);
                    runWidth += advance;
                }
                x += advance;
            }
            Flush();
        }
        else
        {
            double x = text.X;
            int start = 0;
            double startX = x;
            for (int i = 0; i < text.Text.Length;)
            {
                int length = char.IsSurrogatePair(text.Text, i) ? 2 : 1;
                string ch = text.Text.Substring(i, length);
                double advance = text.Font.MeasureText(ch, text.FontSize) + text.CharacterSpacing;
                if (IsUpright(char.ConvertToUtf32(text.Text, i)))
                {
                    if (i > start) pieces.Add((text.Text[start..i], null, startX, x - startX, false));
                    pieces.Add((ch, null, x, advance, true));
                    any = true;
                    start = i + length;
                    startX = x + advance;
                }
                x += advance;
                i += length;
            }
            if (start < text.Text.Length) pieces.Add((text.Text[start..], null, startX, x - startX, false));
        }
        if (!any) return false;

        foreach (var (str, shaped, px, width, upright) in pieces)
        {
            // A piece of a field keeps its field only if it is the whole text; fields are replaced before this runs.
            var piece = text with { X = px, Text = str, Glyphs = shaped, Field = null };
            if (!upright) { output.Add(piece); continue; }
            output.Add(new PushItem(null, 90, px + width / 2, text.Baseline - CenterAboveBaseline * text.FontSize));
            output.Add(piece);
            output.Add(new PopItem());
        }
        return true;
    }
}
