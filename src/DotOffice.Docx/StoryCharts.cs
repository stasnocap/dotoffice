using System.Xml;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Dml;

namespace DotOffice.Docx;

internal sealed partial class StoryReader
{
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>The empty picture a chart run carries in place of an image.</summary>
    private static readonly ImageData ChartPlaceholder = new([]);

    private static bool IsChart(string? graphicType) => graphicType is not null && graphicType.EndsWith("/chart", StringComparison.Ordinal);

    /// <summary>
    /// Reads the chart part a <c>c:chart</c> element points to into a chart of the given extent. A chart that cannot be read or has nothing to
    /// draw yields a content without a chart (its extent stays empty) and a warning; null only for an extent without area.
    /// </summary>
    private (ChartContent Content, double Width, double Height)? LoadChart(string? id, long cx, long cy)
    {
        double width = Units.FromEmu(cx), height = Units.FromEmu(cy);
        if (width <= 0 || height <= 0)
        {
            context.Warn("docx.chart-size", "A chart without an extent was skipped.");
            return null;
        }
        double max = context.Options.MaxDimension;
        if (width > max || height > max)
        {
            double scale = max / Math.Max(width, height);
            (width, height) = (width * scale, height * scale);
            context.Warn("docx.image-size", "An image with excessive dimensions was scaled down.");
        }
        var empty = (new ChartContent(null), width, height);
        if (id is null || !context.Package.GetRelationships(partName).TryGetValue(id, out var rel) || rel.IsExternal || !context.Package.Exists(rel.Target))
        {
            context.Warn("docx.chart-missing", "A chart's part is missing from the file; its space is left empty.");
            return empty;
        }
        XDocument doc;
        try
        {
            using var reader = context.Package.OpenXml(rel.Target);
            doc = XDocument.Load(reader);
        }
        catch (Exception e) when (e is XmlException or InvalidDataException or IOException)
        {
            context.Warn("docx.chart-unreadable", $"Chart '{rel.Target}' could not be read ({e.Message}); its space is left empty.");
            return empty;
        }
        try
        {
            var template = new SheetChart { Anchor = ImageAnchorKind.Absolute, X = 0, Y = 0, Width = width, Height = height };
            var notes = new List<string>();
            var (unsupported, build) = DmlCharts.Prepare(doc, template, notes);
            foreach (var note in notes) context.Warn("docx.chart-approximation", $"Chart '{rel.Target}': {note}.");
            if (build is null)
            {
                context.Warn("docx.chart-unsupported", unsupported is null or "empty"
                    ? $"Chart '{rel.Target}' has no plot to draw; its space is left empty."
                    : $"A {unsupported} chart is not supported; its space is left empty.");
                return empty;
            }
            var chart = build(new WordChartData(context.ChartTheme));
            bool hasValues = chart.Series.Concat(chart.ExtraGroups.SelectMany(g => g.Series)).Any(s => s.Values.Any(v => v is not null));
            if (!hasValues)
            {
                context.Warn("docx.chart-no-cache", $"Chart '{rel.Target}' has no cached values (the embedded workbook is not read); its space is left empty.");
                return empty;
            }
            return (new ChartContent(chart), width, height);
        }
        catch (Exception e) when (e is not OutOfMemoryException)
        {
            context.Warn("docx.chart-unreadable", $"Chart '{rel.Target}' could not be read ({e.Message}); its space is left empty.");
            return empty;
        }
    }

    /// <summary>A chart's data as it reads in a Word document: only its cached values, no cells; defaults as PowerPoint's, with black text.</summary>
    private sealed class WordChartData(DmlTheme theme) : IChartData
    {
        public DmlTheme Theme { get; } = theme;

        public Color? TextColor => Color.Black;

        public bool PresentationTitle => true;

        public bool Date1904 => false;

        public bool FramedChartArea => false;

        public bool WordChartFrame => true;

        public bool FitAxisToHeight => true;

        public IEnumerable<Cell?> Cells(string formula) => [];
    }
}
