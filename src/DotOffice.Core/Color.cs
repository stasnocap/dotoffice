using System.Globalization;

namespace DotOffice.Core;

/// <summary>An opaque sRGB color.</summary>
public readonly record struct Color(byte R, byte G, byte B)
{
    public static Color Black { get; } = new(0, 0, 0);
    public static Color White { get; } = new(255, 255, 255);

    /// <summary>Parses "RRGGBB" (optionally prefixed with '#'); returns null for anything else.</summary>
    public static Color? TryParseHex(string? hex)
    {
        if (hex is null) return null;
        var span = hex.AsSpan().Trim();
        if (span.StartsWith("#")) span = span[1..];
        if (span.Length != 6 || !uint.TryParse(span, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v))
            return null;
        return new Color((byte)(v >> 16), (byte)(v >> 8), (byte)v);
    }

    /// <summary>Linear blend: <paramref name="amount"/> 0 returns this color, 1 returns <paramref name="other"/>.</summary>
    public Color Blend(Color other, double amount)
    {
        amount = Math.Clamp(amount, 0, 1);
        static byte Mix(byte a, byte b, double t) => (byte)Math.Round(a + (b - a) * t);
        return new Color(Mix(R, other.R, amount), Mix(G, other.G, amount), Mix(B, other.B, amount));
    }

    /// <summary>The color with its hue, saturation and lightness (each 0–1) changed by <paramref name="change"/>.</summary>
    public Color WithHsl(Func<double, double, double, (double H, double S, double L)> change)
    {
        double r = R / 255.0, g = G / 255.0, b = B / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), l = (max + min) / 2, h = 0, s = 0;
        if (max - min > 1e-9)
        {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            h = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
            h /= 6;
        }
        (h, s, l) = change(h, s, l);
        l = Math.Clamp(l, 0, 1);
        s = Math.Clamp(s, 0, 1);
        if (s < 1e-9) return new Color((byte)Math.Round(l * 255), (byte)Math.Round(l * 255), (byte)Math.Round(l * 255));
        double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
        static double Channel(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            return t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p;
        }
        return new Color((byte)Math.Round(Channel(p, q, h + 1.0 / 3) * 255), (byte)Math.Round(Channel(p, q, h) * 255), (byte)Math.Round(Channel(p, q, h - 1.0 / 3) * 255));
    }

    public override string ToString() => $"#{R:X2}{G:X2}{B:X2}";
}