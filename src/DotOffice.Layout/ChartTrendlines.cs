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
        var coefficients = Enumerable.Range(0, n).Select(r => m[r, n] / m[r, r]).ToArray();
        return x =>
        {
            double u = (x - mean) / scale, y = 0;
            for (int i = n - 1; i >= 0; i--) y = y * u + coefficients[i];
            return y;
        };
    }
}
