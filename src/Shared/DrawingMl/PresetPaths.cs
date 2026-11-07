using DotOffice.Core;

namespace DotOffice.Dml;

/// <summary>
/// Preset shapes that have no enumerated geometry of their own, built as paths in fractions of the frame (the form custom geometry
/// uses). The guides follow the preset definitions of the Office Open XML specification; shading (the darker or lighter faces of a cube
/// or a bevel) is not drawn, only the outlines.
/// </summary>
internal static class PresetPaths
{
    private const double Pi = Math.PI;

    /// <summary>Collects subpaths given in absolute units and turns them into fractions of the frame.</summary>
    private sealed class Builder(double w, double h)
    {
        public readonly double W = w, H = h;
        public readonly List<ShapePath> Result = [];
        private readonly List<(double X, double Y)> _points = [];
        private (double X, double Y) _current;

        public void Begin(double x, double y) { _points.Clear(); Add(x, y); }
        private void Add(double x, double y) { _current = (x, y); _points.Add((x / W, y / H)); }
        public void Line(double x, double y) => Add(x, y);

        /// <summary>The spec's <c>arcTo</c>: an arc of an ellipse with radii wR and hR from the current point, starting at stDeg and sweeping swDeg (clockwise on screen when positive).</summary>
        public void Arc(double wR, double hR, double stDeg, double swDeg)
        {
            double start = stDeg * Pi / 180, sweep = swDeg * Pi / 180;
            double cx = _current.X - wR * Math.Cos(start), cy = _current.Y - hR * Math.Sin(start);
            int steps = Math.Max(4, (int)Math.Ceiling(Math.Abs(swDeg) / 6));
            for (int k = 1; k <= steps; k++)
            {
                double angle = start + sweep * k / steps;
                Add(cx + wR * Math.Cos(angle), cy + hR * Math.Sin(angle));
            }
        }

        public void Cubic(double x1, double y1, double x2, double y2, double x3, double y3)
        {
            var p0 = _current;
            for (int k = 1; k <= 16; k++)
            {
                double t = k / 16.0, u = 1 - t;
                Add(u * u * u * p0.X + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3,
                    u * u * u * p0.Y + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3);
            }
        }

        public void Quad(double x1, double y1, double x2, double y2)
        {
            var p0 = _current;
            for (int k = 1; k <= 12; k++)
            {
                double t = k / 12.0, u = 1 - t;
                Add(u * u * p0.X + 2 * u * t * x1 + t * t * x2, u * u * p0.Y + 2 * u * t * y1 + t * t * y2);
            }
        }

        public void End(bool closed = true, bool fill = true, bool stroke = true)
        {
            if (_points.Count >= 2) Result.Add(new ShapePath(_points.ToList(), closed, fill, stroke));
            _points.Clear();
        }

        public void Poly(params (double X, double Y)[] points)
        {
            Begin(points[0].X, points[0].Y);
            foreach (var p in points.Skip(1)) Line(p.X, p.Y);
            End();
        }

        public void Stroke(params (double X, double Y)[] points)
        {
            Begin(points[0].X, points[0].Y);
            foreach (var p in points.Skip(1)) Line(p.X, p.Y);
            End(false, false);
        }

        /// <summary>A path of several rings filled together by the even-odd rule: the ring inside another one is a hole.</summary>
        public void Compound(params List<(double X, double Y)>[] rings)
        {
            var fractions = rings.Select(r => r.Select(p => (p.X / W, p.Y / H)).ToList()).ToList();
            Result.Add(new ShapePath(fractions[0], true) { Rings = fractions.Skip(1).Cast<IReadOnlyList<(double X, double Y)>>().ToList() });
        }

        public static List<(double X, double Y)> EllipseRing(double cx, double cy, double rx, double ry)
        {
            var points = new List<(double X, double Y)>();
            for (int k = 0; k < 72; k++) points.Add((cx + rx * Math.Cos(k * Pi / 36), cy + ry * Math.Sin(k * Pi / 36)));
            return points;
        }

        public void Ellipse(double cx, double cy, double rx, double ry, bool fill = true, bool stroke = true)
        {
            Begin(cx + rx, cy);
            Arc(rx, ry, 0, 360);
            End(true, fill, stroke);
        }

        public void MirrorX(int from)
        {
            for (int i = from; i < Result.Count; i++)
                Result[i] = Result[i] with { Points = Result[i].Points.Select(p => (1 - p.X, p.Y)).ToList() };
        }

        /// <summary>Adds polygons given in any units, stretched so that together they fill the frame between the given bounds.</summary>
        public void Fitted(IEnumerable<(List<(double X, double Y)> Points, bool Fill, bool Stroke)> polygons, double? minX = null, double? maxX = null, double? minY = null, double? maxY = null)
        {
            var list = polygons.ToList();
            var all = list.SelectMany(p => p.Points).ToList();
            double x0 = minX ?? all.Min(p => p.X), x1 = maxX ?? all.Max(p => p.X), y0 = minY ?? all.Min(p => p.Y), y1 = maxY ?? all.Max(p => p.Y);
            foreach (var (points, fill, stroke) in list)
                Result.Add(new ShapePath(points.Select(p => ((p.X - x0) / (x1 - x0), (p.Y - y0) / (y1 - y0))).ToList(), true, fill, stroke));
        }
    }

    /// <param name="preset">The preset's name.</param>
    /// <param name="width">The frame's width (any unit; only the proportions matter).</param>
    /// <param name="height">The frame's height.</param>
    /// <param name="raw">The raw value (<c>val</c> of a guide formula, 1/100000 or 1/60000 degree) of a named adjust value, or null when the shape does not state it.</param>
    public static List<ShapePath>? Build(string preset, double width, double height, Func<string, double?> raw)
    {
        double w = width > 0 ? width : 1, h = height > 0 ? height : 1, ss = Math.Min(w, h);
        double A(string name, double fallback, double min = double.NegativeInfinity, double max = double.PositiveInfinity) => Math.Clamp(raw(name) ?? fallback, min, max) / 100_000.0;
        var b = new Builder(w, h);
        switch (preset)
        {
            // Stars: the points on the frame's ellipse, the notches at adj/50000 of that radius.
            case "star4": Star(b, 4, A("adj", 12500, 0, 50000) * 2); break;
            case "star6": Star(b, 6, A("adj", 28868, 0, 50000) * 2); break;
            case "star7": Star(b, 7, A("adj", 34601, 0, 50000) * 2); break;
            case "star8": Star(b, 8, A("adj", 37500, 0, 50000) * 2); break;
            case "star10": Star(b, 10, A("adj", 42533, 0, 50000) * 2); break;
            case "star12": Star(b, 12, A("adj", 37500, 0, 50000) * 2); break;
            case "star16": Star(b, 16, A("adj", 37500, 0, 50000) * 2); break;
            case "star24": Star(b, 24, A("adj", 37500, 0, 50000) * 2); break;
            case "star32": Star(b, 32, A("adj", 37500, 0, 50000) * 2); break;

            // Regular polygons.
            case "heptagon": Regular(b, 7, -90); break;
            case "decagon": Regular(b, 10, -90 + 18); break;
            case "dodecagon": Regular(b, 12, -90 + 15); break;

            // Callouts with a wedge (the tip at adj1 and adj2 of the frame from its center) or with a line.
            case "wedgeRectCallout": Wedge(b, A("adj1", -20833), A("adj2", 62500), 0); break;
            case "wedgeRoundRectCallout": Wedge(b, A("adj1", -20833), A("adj2", 62500), ss * A("adj3", 16667)); break;
            case "wedgeEllipseCallout": EllipseWedge(b, A("adj1", -20833), A("adj2", 62500)); break;
            case "cloudCallout": Cloud(b, A("adj1", -20833), A("adj2", 62500)); break;
            case "borderCallout1": case "callout1": case "accentCallout1": case "accentBorderCallout1":
                b.Poly((0, 0), (w, 0), (w, h), (0, h));
                b.Stroke((w * A("adj2", -8333), h * A("adj1", 18750)), (w * A("adj4", -38333), h * A("adj3", 112500)));
                break;
            case "borderCallout2": case "callout2": case "accentCallout2": case "accentBorderCallout2":
                b.Poly((0, 0), (w, 0), (w, h), (0, h));
                b.Stroke((w * A("adj2", -8333), h * A("adj1", 18750)), (w * A("adj4", -16667), h * A("adj3", 18750)), (w * A("adj6", -46667), h * A("adj5", 112500)));
                break;

            // Flow charts.
            case "flowChartDocument":
                b.Begin(0, 0); b.Line(w, 0); b.Line(w, h * 0.8); b.Cubic(w * 0.75, h * 0.8, w * 0.7, h * 0.95, w * 0.5, h * 0.9); b.Cubic(w * 0.3, h * 0.85, w * 0.25, h * 1.0, 0, h * 0.9); b.End();
                break;
            case "flowChartPredefinedProcess":
                b.Poly((0, 0), (w, 0), (w, h), (0, h));
                b.Stroke((w / 8, 0), (w / 8, h)); b.Stroke((w * 7 / 8, 0), (w * 7 / 8, h));
                break;
            case "flowChartInternalStorage":
                b.Poly((0, 0), (w, 0), (w, h), (0, h));
                b.Stroke((w / 8, 0), (w / 8, h)); b.Stroke((0, h / 8), (w, h / 8));
                break;
            case "flowChartManualInput": b.Poly((0, h / 5), (w, 0), (w, h), (0, h)); break;
            case "flowChartManualOperation": b.Poly((0, 0), (w, 0), (w * 0.8, h), (w * 0.2, h)); break;
            case "flowChartDelay":
                b.Begin(0, 0); b.Line(w / 2, 0); b.Arc(w / 2, h / 2, -90, 180); b.Line(0, h); b.End();
                break;
            case "flowChartMerge": b.Poly((0, 0), (w, 0), (w / 2, h)); break;
            case "flowChartExtract": b.Poly((w / 2, 0), (w, h), (0, h)); break;
            case "flowChartDisplay":
                b.Begin(0, h / 2); b.Line(w / 6, 0); b.Line(w * 5 / 6, 0); b.Arc(w / 6, h / 2, -90, 180); b.Line(w / 6, h); b.End();
                break;
            case "flowChartOffpageConnector": b.Poly((0, 0), (w, 0), (w, h * 0.8), (w / 2, h), (0, h * 0.8)); break;
            case "flowChartPunchedCard": b.Poly((w / 5, 0), (w, 0), (w, h), (0, h), (0, h / 5)); break;
            case "flowChartCollate": b.Poly((0, 0), (w, 0), (0, h), (w, h)); break;
            case "flowChartSort":
                b.Poly((w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)); b.Stroke((0, h / 2), (w, h / 2));
                break;
            case "flowChartOr":
                b.Ellipse(w / 2, h / 2, w / 2, h / 2); b.Stroke((0, h / 2), (w, h / 2)); b.Stroke((w / 2, 0), (w / 2, h));
                break;
            case "flowChartSummingJunction":
            {
                double k = Math.Cos(Pi / 4) / 2;
                b.Ellipse(w / 2, h / 2, w / 2, h / 2);
                b.Stroke((w * (0.5 - k), h * (0.5 - k)), (w * (0.5 + k), h * (0.5 + k)));
                b.Stroke((w * (0.5 + k), h * (0.5 - k)), (w * (0.5 - k), h * (0.5 + k)));
                break;
            }
            case "flowChartMagneticDisk": Can(b, h / 6); break;

            // Brackets and braces (an open stroke for the line, the filled shape without a stroke).
            case "leftBracket": case "rightBracket":
            {
                double y1 = ss * A("adj", 8333, 0, 50000 * h / ss);
                int from = b.Result.Count;
                b.Begin(w, 0); b.Arc(w, y1, 270, -90); b.Line(0, h - y1); b.Arc(w, y1, 180, -90); b.End(true, true, false);
                b.Begin(w, 0); b.Arc(w, y1, 270, -90); b.Line(0, h - y1); b.Arc(w, y1, 180, -90); b.End(false, false, true);
                if (preset == "rightBracket") b.MirrorX(from);
                break;
            }
            case "leftBrace": case "rightBrace":
            {
                double y1 = ss * A("adj1", 8333, 0, 25000 * h / ss), y3 = h * A("adj2", 50000, 0, 100000);
                y1 = Math.Min(y1, Math.Min(y3, h - y3));
                int from = b.Result.Count;
                for (int pass = 0; pass < 2; pass++)
                {
                    b.Begin(w, h); b.Arc(w / 2, y1, 90, 90); b.Line(w / 2, y3 + y1); b.Arc(w / 2, y1, 0, -90); b.Arc(w / 2, y1, 90, -90);
                    b.Line(w / 2, y1); b.Arc(w / 2, y1, 180, 90);
                    if (pass == 0) b.End(true, true, false); else b.End(false, false, true);
                }
                if (preset == "rightBrace") b.MirrorX(from);
                break;
            }

            // Solids.
            case "can": Can(b, ss * A("adj", 25000, 0, 50000 * h / ss) / 2); break;
            case "cube":
            {
                double d = ss * A("adj", 25000, 0, 100000 * h / ss);
                b.Poly((0, d), (w - d, d), (w - d, h), (0, h));
                b.Poly((0, d), (d, 0), (w, 0), (w - d, d));
                b.Poly((w - d, d), (w, 0), (w, h - d), (w - d, h));
                break;
            }
            case "bevel":
            {
                double d = ss * A("adj", 12500, 0, 50000);
                b.Poly((0, 0), (w, 0), (w, h), (0, h));
                b.Poly((d, d), (w - d, d), (w - d, h - d), (d, h - d));
                b.Stroke((0, 0), (d, d)); b.Stroke((w, 0), (w - d, d)); b.Stroke((w, h), (w - d, h - d)); b.Stroke((0, h), (d, h - d));
                break;
            }
            case "foldedCorner":
            {
                double d = ss * A("adj", 16667, 0, 50000);
                b.Poly((0, 0), (w, 0), (w, h - d), (w - d, h), (0, h));
                b.Stroke((w - d, h), (w - d, h - d), (w, h - d));
                break;
            }
            case "cloud": Cloud(b, null, null); break;

            // Symbols.
            case "heart":
                b.Begin(w / 2, h); b.Cubic(w * 0.15, h * 0.72, 0, h * 0.5, 0, h * 0.28); b.Cubic(0, h * 0.1, w * 0.15, 0, w * 0.3, 0);
                b.Cubic(w * 0.4, 0, w * 0.48, h * 0.07, w / 2, h * 0.22);
                b.Cubic(w * 0.52, h * 0.07, w * 0.6, 0, w * 0.7, 0); b.Cubic(w * 0.85, 0, w, h * 0.1, w, h * 0.28); b.Cubic(w, h * 0.5, w * 0.85, h * 0.72, w / 2, h);
                b.End();
                break;
            case "lightningBolt":
            {
                (double X, double Y)[] bolt = [(8472, 0), (12860, 6080), (11050, 6797), (16577, 12007), (14767, 12877), (21600, 21600), (10450, 14630), (12170, 13870), (5600, 8700), (7400, 7900), (0, 3640)];
                b.Poly(bolt.Select(p => (p.X / 21600 * w, p.Y / 21600 * h)).ToArray());
                break;
            }
            case "sun":
            {
                double circle = 0.5 - A("adj", 25000, 12500, 46875) * 0.5 - 0.0; // radius of the disc, as a fraction of the frame
                circle = Math.Max(circle, 0.15);
                for (int k = 0; k < 8; k++)
                {
                    double a = k * Pi / 4, spread = Pi / 14, baseR = circle * 1.12;
                    b.Poly((w / 2 + w / 2 * Math.Cos(a), h / 2 + h / 2 * Math.Sin(a)),
                           (w / 2 + w * baseR * Math.Cos(a + spread), h / 2 + h * baseR * Math.Sin(a + spread)),
                           (w / 2 + w * baseR * Math.Cos(a - spread), h / 2 + h * baseR * Math.Sin(a - spread)));
                }
                b.Ellipse(w / 2, h / 2, w * circle, h * circle);
                break;
            }
            case "moon":
            {
                double x1 = w * A("adj", 50000, 0, 87500);
                b.Begin(w, 0); b.Arc(w, h / 2, 270, -180); b.Arc(w - x1, h / 2, 90, 180); b.End();
                break;
            }
            case "smileyFace":
                b.Ellipse(w / 2, h / 2, w / 2, h / 2);
                b.Ellipse(w * 0.3, h * 0.38, w * 0.05, h * 0.06, false); b.Ellipse(w * 0.7, h * 0.38, w * 0.05, h * 0.06, false);
                b.Begin(w * 0.22, h * 0.62); b.Quad(w / 2, h * 0.96, w * 0.78, h * 0.62); b.End(false, false, true);
                break;
            case "teardrop":
            {
                double k = A("adj", 100000, 0, 200000);
                double tx = w / 2 + w / 2 * k, ty = h / 2 - h / 2 * k;
                b.Begin(0, h / 2); b.Arc(w / 2, h / 2, 180, 90);
                b.Line(tx, ty); b.Line(w, h / 2);
                b.Arc(w / 2, h / 2, 0, 180); b.End();
                break;
            }

            // Arrows and bars.
            case "leftRightArrow":
            {
                double shaft = h * A("adj1", 50000, 0, 100000) / 2, head = Math.Min(ss * A("adj2", 50000, 0, 100000 * w / ss), w / 2);
                b.Poly((0, h / 2), (head, 0), (head, h / 2 - shaft), (w - head, h / 2 - shaft), (w - head, 0), (w, h / 2), (w - head, h), (w - head, h / 2 + shaft), (head, h / 2 + shaft), (head, h));
                break;
            }
            case "upDownArrow":
            {
                double shaft = w * A("adj1", 50000, 0, 100000) / 2, head = Math.Min(ss * A("adj2", 50000, 0, 100000 * h / ss), h / 2);
                b.Poly((w / 2, 0), (w, head), (w / 2 + shaft, head), (w / 2 + shaft, h - head), (w, h - head), (w / 2, h), (0, h - head), (w / 2 - shaft, h - head), (w / 2 - shaft, head), (0, head));
                break;
            }
            case "mathPlus":
            {
                double t = ss * A("adj1", 23520, 0, 73490) / 2;
                b.Poly((w / 2 - t, h * 0.13), (w / 2 + t, h * 0.13), (w / 2 + t, h / 2 - t), (w * 0.87, h / 2 - t), (w * 0.87, h / 2 + t), (w / 2 + t, h / 2 + t), (w / 2 + t, h * 0.87), (w / 2 - t, h * 0.87), (w / 2 - t, h / 2 + t), (w * 0.13, h / 2 + t), (w * 0.13, h / 2 - t), (w / 2 - t, h / 2 - t));
                break;
            }
            case "mathMinus":
            {
                double t = h * A("adj1", 23520, 0, 100000) / 2, half = w * A("adj2", 73490, 0, 100000) / 2;
                b.Poly((w / 2 - half, h / 2 - t), (w / 2 + half, h / 2 - t), (w / 2 + half, h / 2 + t), (w / 2 - half, h / 2 + t));
                break;
            }
            case "mathEqual":
            {
                double t = h * A("adj1", 23520, 0, 36745) / 2, gap = h * A("adj2", 11760, 0, 100000) / 2, half = w * 73490 / 200000.0;
                b.Poly((w / 2 - half, h / 2 - gap - t * 2), (w / 2 + half, h / 2 - gap - t * 2), (w / 2 + half, h / 2 - gap), (w / 2 - half, h / 2 - gap));
                b.Poly((w / 2 - half, h / 2 + gap), (w / 2 + half, h / 2 + gap), (w / 2 + half, h / 2 + gap + t * 2), (w / 2 - half, h / 2 + gap + t * 2));
                break;
            }

            // Rectangles with cut or rounded corners.
            case "snip1Rect": { double d = ss * A("adj", 16667, 0, 50000); b.Poly((0, 0), (w - d, 0), (w, d), (w, h), (0, h)); break; }
            case "snip2SameRect":
            {
                double d1 = ss * A("adj1", 16667, 0, 50000), d2 = ss * A("adj2", 0, 0, 50000);
                b.Poly((d1, 0), (w - d1, 0), (w, d1), (w, h - d2), (w - d2, h), (d2, h), (0, h - d2), (0, d1));
                break;
            }
            case "snip2DiagRect":
            {
                double d1 = ss * A("adj1", 0, 0, 50000), d2 = ss * A("adj2", 16667, 0, 50000);
                b.Poly((d1, 0), (w - d2, 0), (w, d2), (w, h - d1), (w - d1, h), (d2, h), (0, h - d2), (0, d1));
                break;
            }
            case "round1Rect":
            {
                double r = ss * A("adj", 16667, 0, 50000);
                b.Begin(0, 0); b.Line(w - r, 0); b.Arc(r, r, 270, 90); b.Line(w, h); b.Line(0, h); b.End();
                break;
            }
            case "round2SameRect":
            {
                double r1 = ss * A("adj1", 16667, 0, 50000), r2 = ss * A("adj2", 0, 0, 50000);
                b.Begin(r1, 0); b.Line(w - r1, 0); b.Arc(r1, r1, 270, 90); b.Line(w, h - r2); b.Arc(r2, r2, 0, 90); b.Line(r2, h); b.Arc(r2, r2, 90, 90); b.Line(0, r1); b.Arc(r1, r1, 180, 90); b.End();
                break;
            }

            // Shapes with a hole.
            case "donut":
            {
                double t = ss * A("adj", 25000, 0, 50000);
                b.Compound(Builder.EllipseRing(w / 2, h / 2, w / 2, h / 2), Builder.EllipseRing(w / 2, h / 2, Math.Max(w / 2 - t, 0.01), Math.Max(h / 2 - t, 0.01)));
                break;
            }
            case "frame":
            {
                double t = ss * A("adj1", 12500, 0, 50000);
                b.Compound([(0, 0), (w, 0), (w, h), (0, h)], [(t, t), (w - t, t), (w - t, h - t), (t, h - t)]);
                break;
            }
            case "noSmoking":
            {
                double t = ss * A("adj", 18750, 0, 50000 * h / ss);
                double rx = Math.Max(w / 2 - t, 0.01), ry = Math.Max(h / 2 - t, 0.01);
                b.Compound(Builder.EllipseRing(w / 2, h / 2, w / 2, h / 2), Builder.EllipseRing(w / 2, h / 2, rx, ry));
                // The slash: a band across the inner circle from the top left to the bottom right, as thick as the ring.
                double phi = Math.Max(Math.Asin(Math.Min(t / 2 / Math.Min(rx, ry), 1)), 0.01);
                var slash = new List<(double X, double Y)>();
                for (double a = Pi / 4 - phi; a <= Pi / 4 + phi + 1e-9; a += phi / 6) slash.Add((w / 2 + rx * Math.Cos(a), h / 2 + ry * Math.Sin(a)));
                for (double a = 5 * Pi / 4 - phi; a <= 5 * Pi / 4 + phi + 1e-9; a += phi / 6) slash.Add((w / 2 + rx * Math.Cos(a), h / 2 + ry * Math.Sin(a)));
                b.Poly(slash.ToArray());
                break;
            }

            // Pieces of an ellipse.
            case "pie":
            {
                double st = raw("adj1") ?? 0, en = raw("adj2") ?? 16_200_000;
                double sw = ((en - st) % 21_600_000 + 21_600_000) % 21_600_000 / 60_000;
                b.Begin(w / 2, h / 2); b.Line(w / 2 + w / 2 * Math.Cos(st / 60_000 * Pi / 180), h / 2 + h / 2 * Math.Sin(st / 60_000 * Pi / 180));
                b.Arc(w / 2, h / 2, st / 60_000, sw); b.End();
                break;
            }
            case "chord":
            {
                double st = raw("adj1") ?? 2_700_000, en = raw("adj2") ?? 16_200_000;
                double sw = ((en - st) % 21_600_000 + 21_600_000) % 21_600_000 / 60_000;
                b.Begin(w / 2 + w / 2 * Math.Cos(st / 60_000 * Pi / 180), h / 2 + h / 2 * Math.Sin(st / 60_000 * Pi / 180));
                b.Arc(w / 2, h / 2, st / 60_000, sw); b.End();
                break;
            }
            case "arc":
            {
                double st = raw("adj1") ?? 16_200_000, en = raw("adj2") ?? 0;
                double sw = ((en - st) % 21_600_000 + 21_600_000) % 21_600_000 / 60_000;
                double sx = w / 2 + w / 2 * Math.Cos(st / 60_000 * Pi / 180), sy = h / 2 + h / 2 * Math.Sin(st / 60_000 * Pi / 180);
                b.Begin(w / 2, h / 2); b.Line(sx, sy); b.Arc(w / 2, h / 2, st / 60_000, sw); b.End(true, true, false);
                b.Begin(sx, sy); b.Arc(w / 2, h / 2, st / 60_000, sw); b.End(false, false, true);
                break;
            }
            case "blockArc":
            {
                double st = raw("adj1") ?? 10_800_000, en = raw("adj2") ?? 0, t = ss * A("adj3", 25000, 0, 50000);
                double sw = ((en - st) % 21_600_000 + 21_600_000) % 21_600_000 / 60_000;
                double a0 = st / 60_000 * Pi / 180, a1 = en / 60_000 * Pi / 180;
                b.Begin(w / 2 + w / 2 * Math.Cos(a0), h / 2 + h / 2 * Math.Sin(a0)); b.Arc(w / 2, h / 2, st / 60_000, sw);
                b.Line(w / 2 + (w / 2 - t) * Math.Cos(a1), h / 2 + (h / 2 - t) * Math.Sin(a1)); b.Arc(w / 2 - t, h / 2 - t, en / 60_000, -sw); b.End();
                break;
            }
            default: return null;
        }
        return b.Result;
    }

    private static void Star(Builder b, int n, double ratio)
    {
        var points = new List<(double X, double Y)>();
        for (int k = 0; k < 2 * n; k++)
        {
            double angle = (-90 + 180.0 * k / n) * Pi / 180, r = k % 2 == 0 ? 1 : ratio;
            points.Add((r * Math.Cos(angle), r * Math.Sin(angle)));
        }
        b.Fitted([(points, true, true)]);
    }

    private static void Regular(Builder b, int n, double startDeg)
    {
        var points = Enumerable.Range(0, n).Select(k => (X: Math.Cos((startDeg + 360.0 * k / n) * Pi / 180), Y: Math.Sin((startDeg + 360.0 * k / n) * Pi / 180))).ToList();
        b.Fitted([(points, true, true)]);
    }

    /// <summary>A rectangle (rounded when r &gt; 0) with a wedge that points at (dx, dy) of the frame from its center, leaving the side the tip is nearest.</summary>
    private static void Wedge(Builder b, double dx, double dy, double r)
    {
        double w = b.W, h = b.H, tx = w / 2 + w * dx, ty = h / 2 + h * dy;
        r = Math.Min(r, Math.Min(w, h) / 2);
        bool horizontal = Math.Abs(dx) > Math.Abs(dy);
        int side = horizontal ? (dx > 0 ? 1 : 3) : (dy > 0 ? 2 : 0); // 0 top, 1 right, 2 bottom, 3 left
        (double A, double B) Span(double length, bool low) => low ? (length / 6, length * 5 / 12) : (length * 7 / 12, length * 5 / 6);
        b.Begin(r, 0);
        if (side == 0) { var (a, c) = Span(w, dx < 0); b.Line(a, 0); b.Line(tx, ty); b.Line(c, 0); }
        b.Line(w - r, 0); if (r > 0) b.Arc(r, r, 270, 90);
        if (side == 1) { var (a, c) = Span(h, dy < 0); b.Line(w, a); b.Line(tx, ty); b.Line(w, c); }
        b.Line(w, h - r); if (r > 0) b.Arc(r, r, 0, 90);
        if (side == 2) { var (a, c) = Span(w, dx < 0); b.Line(c, h); b.Line(tx, ty); b.Line(a, h); }
        b.Line(r, h); if (r > 0) b.Arc(r, r, 90, 90);
        if (side == 3) { var (a, c) = Span(h, dy < 0); b.Line(0, c); b.Line(tx, ty); b.Line(0, a); }
        b.Line(0, r); if (r > 0) b.Arc(r, r, 180, 90);
        b.End();
    }

    private static void EllipseWedge(Builder b, double dx, double dy)
    {
        double w = b.W, h = b.H, tx = w / 2 + w * dx, ty = h / 2 + h * dy;
        double theta = Math.Atan2(dy * 2, dx * 2) * 180 / Pi, gap = 18;
        double sx = w / 2 + w / 2 * Math.Cos((theta + gap) * Pi / 180), sy = h / 2 + h / 2 * Math.Sin((theta + gap) * Pi / 180);
        b.Begin(sx, sy); b.Arc(w / 2, h / 2, theta + gap, 360 - 2 * gap); b.Line(tx, ty); b.End();
    }

    /// <summary>A cloud: scallops (half circles) over the chords of a ring of points; with a tip, the cloud callout's trail of circles toward it.</summary>
    private static void Cloud(Builder b, double? tipDx, double? tipDy)
    {
        const int Bumps = 9;
        var ring = Enumerable.Range(0, Bumps).Select(k => (X: 0.82 * Math.Cos(2 * Pi * k / Bumps - Pi / 2), Y: 0.82 * Math.Sin(2 * Pi * k / Bumps - Pi / 2))).ToList();
        var outline = new List<(double X, double Y)>();
        for (int k = 0; k < Bumps; k++)
        {
            var p = ring[k]; var q = ring[(k + 1) % Bumps];
            double mx = (p.X + q.X) / 2, my = (p.Y + q.Y) / 2, radius = Math.Sqrt((p.X - q.X) * (p.X - q.X) + (p.Y - q.Y) * (p.Y - q.Y)) / 2;
            double start = Math.Atan2(p.Y - my, p.X - mx);
            for (int s = 0; s < 12; s++) { double a = start + Pi * s / 12; outline.Add((mx + radius * Math.Cos(a), my + radius * Math.Sin(a))); }
        }
        double x0 = outline.Min(p => p.X), x1 = outline.Max(p => p.X), y0 = outline.Min(p => p.Y), y1 = outline.Max(p => p.Y);
        if (tipDx is not { } dx || tipDy is not { } dy)
        {
            b.Fitted([(outline, true, true)]);
            return;
        }
        // The cloud sits in the unit frame and the trail reaches (0.5 + dx, 0.5 + dy); the whole is then stretched to the frame.
        (double X, double Y) Unit((double X, double Y) p) => ((p.X - x0) / (x1 - x0), (p.Y - y0) / (y1 - y0));
        var cloud = outline.Select(Unit).ToList();
        double tx = 0.5 + dx, ty = 0.5 + dy;
        double fx0 = Math.Min(0, tx), fx1 = Math.Max(1, tx), fy0 = Math.Min(0, ty), fy1 = Math.Max(1, ty);
        var parts = new List<(List<(double X, double Y)>, bool, bool)> { (cloud, true, true) };
        double[] at = [0.6, 0.8, 0.95], size = [0.07, 0.05, 0.035];
        for (int k = 0; k < 3; k++)
        {
            double cx = 0.5 + (tx - 0.5) * at[k], cy = 0.5 + (ty - 0.5) * at[k], r = size[k] * Math.Max(fx1 - fx0, fy1 - fy0);
            parts.Add((Enumerable.Range(0, 24).Select(s => (cx + r * Math.Cos(s * Pi / 12), cy + r * Math.Sin(s * Pi / 12))).ToList(), true, true));
        }
        b.Fitted(parts, fx0, fx1, fy0, fy1);
    }

    /// <summary>A cylinder: the body with a half-ellipse of half-height <paramref name="cap"/> at the top and bottom, and the top's front edge.</summary>
    private static void Can(Builder b, double cap)
    {
        double w = b.W, h = b.H;
        cap = Math.Min(cap, h / 2);
        b.Begin(0, cap); b.Arc(w / 2, cap, 180, 180); b.Line(w, h - cap); b.Arc(w / 2, cap, 0, 180); b.End();
        b.Begin(0, cap); b.Arc(w / 2, cap, 180, -180); b.End(false, false, true);
    }
}
