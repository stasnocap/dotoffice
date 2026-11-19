using System.Globalization;
using System.Text;

namespace DotOffice.Docx;

/// <summary>Formats list and note numbers in Word's number formats (w:numFmt names).</summary>
internal static class ListNumbers
{
    public static string Format(int value, string format) => format switch
    {
        "decimalZero" => value < 10 && value >= 0 ? "0" + value : value.ToString(CultureInfo.InvariantCulture),
        "upperRoman" => Roman(value),
        "lowerRoman" => Roman(value).ToLowerInvariant(),
        "upperLetter" => Letters(value),
        "lowerLetter" => Letters(value).ToLowerInvariant(),
        "ordinal" => value + OrdinalSuffix(value),
        "none" => "",
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    private static string Roman(int value)
    {
        if (value <= 0 || value >= 4000) return value.ToString(CultureInfo.InvariantCulture);
        ReadOnlySpan<int> numbers = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        string[] symbols = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
        var sb = new StringBuilder();
        for (int i = 0; i < numbers.Length; i++)
            while (value >= numbers[i]) { sb.Append(symbols[i]); value -= numbers[i]; }
        return sb.ToString();
    }

    /// <summary>Word letter numbering: A..Z, then AA..ZZ, AAA… (the letter repeats).</summary>
    private static string Letters(int value)
    {
        if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
        int repeat = (value - 1) / 26 + 1;
        return new string((char)('A' + (value - 1) % 26), Math.Min(repeat, 50));
    }

    private static string OrdinalSuffix(int value) => (value % 100) switch
    {
        11 or 12 or 13 => "th",
        _ => (value % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" },
    };
}
