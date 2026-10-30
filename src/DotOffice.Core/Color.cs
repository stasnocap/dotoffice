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

    public override string ToString() => $"#{R:X2}{G:X2}{B:X2}";
}