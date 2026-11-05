using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>Least-squares fits for chart trendlines.</summary>
internal static class ChartTrendlines
{
    /// <summary>
    /// The points of the trendline over the data range [<paramref name="first"/>, <paramref name="last"/>] widened by the trendline's
    /// backward and forward periods: a curve sampled along x, or for a moving average one point per data point.
    /// Returns nothing when the data does not allow the fit (too few points, or non-positive values for a logarithmic fit).
    /// </summary>
    public static List<(double X, double Y)> Fit(ChartTrendline trend, IReadOnlyList<(double X, double Y)> data, double first, double last)
    {
        var result = new List<(double X, double Y)>();
        var points = data.OrderBy(p => p.X).ToList();
        if (trend.Kind == TrendlineKind.MovingAverage)
        {
            int period = Math.Min(trend.Period, points.Count);
            for (int i = period - 1; period >= 2 && i < points.Count; i++)
                result.Add((points[i].X, points.Skip(i - period + 1).Take(period).Average(p => p.Y)));
            return result;
        }
        if (points.Count < 2) return result;
        double x0 = Math.Min(first, points[0].X) - trend.Backward, x1 = Math.Max(last, points[^1].X) + trend.Forward;
        if (trend.Kind == TrendlineKind.Polynomial ? points.Count <= 1 : x1 <= x0) return result;

        Func<double, double>? f = trend.Kind switch
        {
            TrendlineKind.Linear => Polynomial(points, 1),
            TrendlineKind.Polynomial => Polynomial(points, Math.Min(trend.Order, points.Count - 1)),
            TrendlineKind.Exponential => Transformed(points, p => p.X, p => p.Y > 0 ? Math.Log(p.Y) : double.NaN, (a, b) => x => Math.Exp(a + b * x)),
            TrendlineKind.Logarithmic => Transformed(points, p => p.X > 0 ? Math.Log(p.X) : double.NaN, p => p.Y, (a, b) => x => a + b * Math.Log(x)),
            TrendlineKind.Power => Transformed(points, p => p.X > 0 ? Math.Log(p.X) : double.NaN, p => p.Y > 0 ? Math.Log(p.Y) : double.NaN, (a, b) => x => Math.Exp(a) * Math.Pow(x, b)),
            _ => null,
        };
        if (f is null) return result;
        bool positiveOnly = trend.Kind is TrendlineKind.Logarithmic or TrendlineKind.Power;
        if (positiveOnly && x0 <= 0) x0 = points.Where(p => p.X > 0).Select(p => p.X).DefaultIfEmpty(1).Min() / 2;
        int steps = trend.Kind is TrendlineKind.Linear ? 1 : 48;
        for (int k = 0; k <= steps; k++)
        {
            double x = x0 + (x1 - x0) * k / steps, y = f(x);
            if (double.IsFinite(y)) result.Add((x, y));
        }
        return result;
    }

    /// <summary>The line fitted to values transformed by <paramref name="tx"/> and <paramref name="ty"/>; points that cannot be transformed are skipped.</summary>
    private static Func<double, double>? Transformed(List<(double X, double Y)> points, Func<(double X, double Y), double> tx, Func<(double X, double Y), double> ty,
        Func<double, double, Func<double, double>> build)
    {
        var usable = points.Select(p => (X: tx(p), Y: ty(p))).Where(p => double.IsFinite(p.X) && double.IsFinite(p.Y)).ToList();
        if (usable.Count < 2 || Polynomial(usable, 1) is not { } line) return null;
        double a = line(0), b = line(1) - a;
        return build(a, b);
    }

    /// <summary>The least-squares polynomial of the given degree (normal equations on x centered at its mean, solved by Gaussian elimination).</summary>
    private static Func<double, double>? Polynomial(List<(double X, double Y)> points, int degree)
    {
        if (PolyFit(points, degree) is not var (coefficients, mean, scale)) return null;
        int n = degree + 1;
        return x =>
        {
            double u = (x - mean) / scale, y = 0;
            for (int i = n - 1; i >= 0; i--) y = y * u + coefficients[i];
            return y;
        };
    }

    /// <summary>The coefficients in u = (x − mean) / scale, lowest degree first, with the mean and the scale.</summary>
    private static (double[] Coefficients, double Mean, double Scale)? PolyFit(List<(double X, double Y)> points, int degree)
    {
        if (points.Count <= degree || degree < 1) return null;
        double mean = points.Average(p => p.X), scale = Math.Max(points.Max(p => Math.Abs(p.X - mean)), 1e-12);
        int n = degree + 1;
        var m = new double[n, n + 1];
        foreach (var (x, y) in points)
        {
            double u = (x - mean) / scale;
            for (int r = 0; r < n; r++)
            {
                for (int c = 0; c < n; c++) m[r, c] += Math.Pow(u, r + c);
                m[r, n] += y * Math.Pow(u, r);
            }
        }
        for (int col = 0; col < n; col++)
        {
            int pivot = col;
            for (int r = col + 1; r < n; r++) if (Math.Abs(m[r, col]) > Math.Abs(m[pivot, col])) pivot = r;
            if (Math.Abs(m[pivot, col]) < 1e-12) return null;
            if (pivot != col) for (int c = 0; c <= n; c++) (m[col, c], m[pivot, c]) = (m[pivot, c], m[col, c]);
            for (int r = 0; r < n; r++)
            {
                if (r == col) continue;
                double factor = m[r, col] / m[col, col];
                for (int c = col; c <= n; c++) m[r, c] -= factor * m[col, c];
            }
        }
        return (Enumerable.Range(0, n).Select(r => m[r, n] / m[r, r]).ToArray(), mean, scale);
    }

    /// <summary>The same polynomial in x itself (lowest degree first), expanded from the centered form.</summary>
    private static double[]? RawPolynomial(List<(double X, double Y)> points, int degree)
    {
        if (PolyFit(points, degree) is not var (a, mean, scale)) return null;
        int n = degree + 1;
        var raw = new double[n];
        // a_k u^k = a_k / s^k (x − m)^k, and (x − m)^k = Σ C(k, j) x^j (−m)^(k−j).
        for (int k = 0; k < n; k++)
        {
            double factor = a[k] / Math.Pow(scale, k), binomial = 1;
            for (int j = 0; j <= k; j++)
            {
                raw[j] += factor * binomial * Math.Pow(-mean, k - j);
                binomial = binomial * (k - j) / (j + 1);
            }
        }
        return raw;
    }

    /// <summary>
    /// The label Excel prints for a trendline: its equation and R², fitted like the line itself (R² of the exponential, logarithmic and
    /// power fits is that of the transformed data). Null for a moving average or when the data does not allow the fit.
    /// </summary>
    public static (string Equation, double RSquared)? Describe(ChartTrendline trend, IReadOnlyList<(double X, double Y)> data)
    {
        var points = data.OrderBy(p => p.X).ToList();
        if (trend.Kind == TrendlineKind.MovingAverage || points.Count < 2) return null;
        double[]? coefficients;
        List<(double X, double Y)> fitted;       // the data in the space where the fit is a polynomial
        switch (trend.Kind)
        {
            case TrendlineKind.Linear or TrendlineKind.Polynomial:
                fitted = points;
                break;
            case TrendlineKind.Exponential:
                fitted = points.Where(p => p.Y > 0).Select(p => (p.X, Math.Log(p.Y))).ToList();
                break;
            case TrendlineKind.Logarithmic:
                fitted = points.Where(p => p.X > 0).Select(p => (Math.Log(p.X), p.Y)).ToList();
                break;
            default:
                fitted = points.Where(p => p.X > 0 && p.Y > 0).Select(p => (Math.Log(p.X), Math.Log(p.Y))).ToList();
                break;
        }
        int degree = trend.Kind == TrendlineKind.Polynomial ? Math.Min(trend.Order, fitted.Count - 1) : 1;
        if (fitted.Count < 2 || degree < 1 || (coefficients = RawPolynomial(fitted, degree)) is null) return null;
        double Predict(double x) { double y = 0; for (int i = coefficients.Length - 1; i >= 0; i--) y = y * x + coefficients[i]; return y; }
        double mean = fitted.Average(p => p.Y), total = fitted.Sum(p => (p.Y - mean) * (p.Y - mean)), residual = fitted.Sum(p => (p.Y - Predict(p.X)) * (p.Y - Predict(p.X)));
        double r2 = total < 1e-300 ? 1 : Math.Clamp(1 - residual / total, 0, 1);

        string equation = trend.Kind switch
        {
            TrendlineKind.Exponential => $"y = {Number(Math.Exp(coefficients[0]))}e^{Number(coefficients[1])}x",
            TrendlineKind.Logarithmic => $"y = {Term(coefficients[1], "ln(x)", first: true)}{Sign(coefficients[0])}",
            TrendlineKind.Power => $"y = {Number(Math.Exp(coefficients[0]))}x^{Number(coefficients[1])}",
            _ => "y = " + Polynomial(coefficients),
        };
        return (equation, r2);
    }

    private static string Polynomial(double[] c)
    {
        var text = new System.Text.StringBuilder();
        double largest = c.Max(Math.Abs);
        for (int k = c.Length - 1; k >= 0; k--)
        {
            if (Math.Abs(c[k]) < largest * 1e-9 && k > 0) continue; // rounding noise of the fit, not a term
            string variable = k switch { 0 => "", 1 => "x", 2 => "x²", 3 => "x³", _ => $"x^{k}" };
            if (text.Length == 0) text.Append(Term(c[k], variable, first: true));
            else text.Append(Term(c[k], variable, first: false));
        }
        return text.Length == 0 ? "0" : text.ToString();
    }

    private static string Term(double value, string variable, bool first)
    {
        if (variable.Length == 0) return first ? Number(value) : Sign(value);
        string magnitude = Math.Abs(Math.Abs(value) - 1) < 5e-5 ? "" : Number(Math.Abs(value));
        string sign = value < 0 ? (first ? "-" : " - ") : first ? "" : " + ";
        return sign + magnitude + variable;
    }

    private static string Sign(double value) => value < 0 ? " - " + Number(-value) : " + " + Number(value);

    /// <summary>Five significant digits, without a trailing zero and without an exponent unless the number is tiny or huge.</summary>
    internal static string Number(double value)
    {
        if (value == 0 || !double.IsFinite(value)) return "0";
        double magnitude = Math.Abs(value);
        if (magnitude < 1e-4 || magnitude >= 1e10) return value.ToString("0.####E+00", System.Globalization.CultureInfo.InvariantCulture);
        int digits = Math.Clamp(4 - (int)Math.Floor(Math.Log10(magnitude)), 0, 15);
        return Math.Round(value, digits).ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture);
    }
}
