using System.Globalization;
using System.Text;

namespace DotOffice.Xlsx;

/// <summary>Formats cell numbers the way Excel displays them.</summary>
internal static class NumberFormatter
{
    /// <summary>
    /// Excel's "General" format: up to 11 significant digits in plain notation; values whose exponent falls outside
    /// 1E-4…1E+10 are shown in scientific notation with up to five decimals (<c>1.23457E+11</c>).
    /// </summary>
    public static string General(double value)
    {
        if (double.IsNaN(value)) return "#NUM!";
        if (double.IsInfinity(value)) return "#NUM!";
        if (value == 0) return "0";

        // 11 significant digits: d.dddddddddd E+xxx
        var e = Math.Abs(value).ToString("E10", CultureInfo.InvariantCulture);
        int ePos = e.IndexOf('E');
        var digits = e[..ePos].Replace(".", "");
        int exponent = int.Parse(e.AsSpan(ePos + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        var trimmed = digits.TrimEnd('0');
        if (trimmed.Length == 0) trimmed = "0";

        var sb = new StringBuilder();
        if (value < 0) sb.Append('-');
        if (exponent is >= -4 and <= 10)
        {
            if (exponent >= 0)
            {
                sb.Append(trimmed.Length <= exponent + 1 ? trimmed.PadRight(exponent + 1, '0') : trimmed[..(exponent + 1)]);
                if (trimmed.Length > exponent + 1) sb.Append('.').Append(trimmed.AsSpan(exponent + 1));
            }
            else
            {
                sb.Append("0.").Append('0', -exponent - 1).Append(trimmed);
            }
        }
        else
        {
            // Scientific: mantissa rounded to six significant digits.
            var sci = Math.Abs(value).ToString("E5", CultureInfo.InvariantCulture);
            int p = sci.IndexOf('E');
            var mantissa = sci[..p].TrimEnd('0').TrimEnd('.');
            int exp = int.Parse(sci.AsSpan(p + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            sb.Append(mantissa).Append('E').Append(exp < 0 ? '-' : '+').Append(Math.Abs(exp).ToString("00", CultureInfo.InvariantCulture));
        }
        return sb.ToString();
    }
}
