using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Dml;
using static DotOffice.Dml.DrawingMl;

namespace DotOffice.Xlsx;

/// <summary>Reads the pictures and shapes of a worksheet's drawing part (SpreadsheetML Drawing, <c>xdr:wsDr</c>).</summary>
internal static class XlsxDrawings
{
    private const double EmuPerPoint = 12700;

    /// <summary>What a drawing contained that is not rendered, and what is rendered only approximately.</summary>
    internal sealed record Skipped(int Charts, int Shapes, int Other, IReadOnlyList<string> Notes, IReadOnlyList<string>? ChartKinds = null)
    {
        public bool Any => Charts + Shapes + Other > 0;
    }

    /// <summary>
    /// Adds the pictures and shapes of the drawing part to <paramref name="sheet"/> and reports what was left out. Objects that are hidden or
    /// excluded from printing (<c>fPrintsWithSheet="0"</c>, as Excel's "Print object" option writes it) are not printed.
    /// </summary>
    public static Skipped Read(XlsxContext context, Worksheet sheet, string drawingPart)
    {
        XDocument doc;
        using (var reader = context.Package.OpenXml(drawingPart)) doc = XDocument.Load(reader);
        if (doc.Root is not { } root || root.Name.LocalName != "wsDr") return new Skipped(0, 0, 1, []);
        var state = new State(context, sheet, context.Package.GetRelationships(drawingPart));

        foreach (var anchor in root.Elements())
        {
            var kind = anchor.Name.LocalName switch
            {
                "twoCellAnchor" => ImageAnchorKind.TwoCell,
                "oneCellAnchor" => ImageAnchorKind.OneCell,
                "absoluteAnchor" => ImageAnchorKind.Absolute,
                _ => (ImageAnchorKind?)null,
            };
            if (kind is null) continue;
            if ((string?)Child(anchor, "clientData")?.Attribute("fPrintsWithSheet") is "0" or "false") continue;

            var frame = new AnchorFrame(kind.Value, Point(Child(anchor, "from")), Point(Child(anchor, "to")), Number(Child(anchor, "pos"), "x") / EmuPerPoint,
                Number(Child(anchor, "pos"), "y") / EmuPerPoint, Number(Child(anchor, "ext"), "cx") / EmuPerPoint, Number(Child(anchor, "ext"), "cy") / EmuPerPoint,
                Child(anchor, "ext") is not null);
            foreach (var content in anchor.Elements())
            {
                switch (content.Name.LocalName)
                {
                    case "pic" or "sp" or "cxnSp" or "grpSp":
                        ReadObject(state, frame, FrameFraction.Whole, content);
                        break;
                    case "graphicFrame": ReadChart(state, frame, FrameFraction.Whole, content); break;
                    case "from" or "to" or "pos" or "ext" or "clientData": break;
                    default: state.Other++; break; // AlternateContent (slicers, …), contentPart
                }
            }
        }
        return new Skipped(state.Charts, state.Shapes, state.Other, state.Notes, state.ChartKinds);
    }

    /// <summary>The anchor of a top-level object: where its frame is, and whether the anchor itself states the size (<c>ext</c>).</summary>
    private sealed record AnchorFrame(ImageAnchorKind Kind, AnchorPoint From, AnchorPoint To, double X, double Y, double Width, double Height, bool HasExtent);

    private sealed class State(XlsxContext context, Worksheet sheet, IReadOnlyDictionary<string, Opc.Relationship> rels) : IDmlHost
    {
        public DmlTheme Theme => Context.Theme;

        public ImageData? Image(string relationshipId, out bool unreadable)
        {
            unreadable = false;
            if (!Relationships.TryGetValue(relationshipId, out var rel) || rel.IsExternal || !Context.Package.Exists(rel.Target)) return null;
            try { return Context.GetImage(rel.Target); }
            catch (InvalidDataException)
            {
                unreadable = true;
                return null;
            }
        }

        public XlsxContext Context { get; } = context;
        public Worksheet Sheet { get; } = sheet;
        public IReadOnlyDictionary<string, Opc.Relationship> Relationships { get; } = rels;
        public int Charts, Shapes, Other, ZOrder;
        public List<string> Notes { get; } = [];
        public List<string> ChartKinds { get; } = [];

        /// <summary>The rotated groups being read, outermost first.</summary>
        public List<GroupTurn> Turns { get; } = [];

        /// <summary>An object's own rotation plus the turns of its groups.</summary>
        public double Turned(double own) => own + Turns.Sum(t => t.Rotation);

        public void Note(string note)
        {
            if (!Notes.Contains(note)) Notes.Add(note);
        }
    }

    private static void ReadObject(State state, AnchorFrame anchor, FrameFraction frame, XElement element)
    {
        switch (element.Name.LocalName)
        {
            case "pic": ReadPicture(state, anchor, frame, element); break;
            case "sp" or "cxnSp": ReadShape(state, anchor, frame, element); break;
            case "grpSp": ReadGroup(state, anchor, frame, element); break;
            case "graphicFrame": ReadChart(state, anchor, frame, element); break;
        }
    }

    /// <summary>Reads a chart frame; charts of kinds that are not drawn are counted and named for the warning.</summary>
    private static void ReadChart(State state, AnchorFrame anchor, FrameFraction frame, XElement graphicFrame)
    {
        var context = state.Context;
        if (Child(Child(graphicFrame, "nvGraphicFramePr"), "cNvPr") is { } p && (string?)p.Attribute("hidden") is "1" or "true") return;
        var chartElement = graphicFrame.Descendants().FirstOrDefault(e => e.Name.LocalName == "chart" && e.Attributes().Any(a => a.Name.LocalName == "id"));
        var id = chartElement?.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
        if (id is null || !state.Relationships.TryGetValue(id, out var rel) || rel.IsExternal || !context.Package.Exists(rel.Target))
        {
            // Not a chart (a table or a diagram) or a chart whose part is missing.
            state.Charts++;
            return;
        }
        XDocument doc;
        try
        {
            using var reader = context.Package.OpenXml(rel.Target);
            doc = XDocument.Load(reader);
        }
        catch (System.Xml.XmlException)
        {
            state.Charts++;
            return;
        }
        var (kind, from, to, x, y, width, height) = Placement(anchor, Child(graphicFrame, "xfrm"), frame);
        var template = new SheetChart { Anchor = kind, From = from, To = to, X = x, Y = y, Width = width, Height = height, Frame = frame, GroupTurns = state.Turns.ToArray(), Rotation = state.Turned(0), ZOrder = state.ZOrder++ };
        var (unsupported, build) = DmlCharts.Prepare(doc, template, state.Notes);
        if (build is null)
        {
            state.Charts++;
            state.ChartKinds.Add(unsupported ?? "unknown");
            return;
        }
        var owner = state.Sheet;
        context.DeferChart(owner, sheets => build(new ChartCells(context, owner, sheets)));
    }

    /// <summary>
    /// Reads a group by placing its children inside the group's frame: a child's own coordinates are mapped from the group's child space
    /// (<c>chOff</c>, <c>chExt</c>) onto the group's frame. A rotated group swings its children about its center and turns them with it.
    /// </summary>
    private static void ReadGroup(State state, AnchorFrame anchor, FrameFraction frame, XElement group)
    {
        if (Child(Child(group, "nvGrpSpPr"), "cNvPr") is { } p && (string?)p.Attribute("hidden") is "1" or "true") return;
        var xfrm = Child(Child(group, "grpSpPr"), "xfrm");
        double turn = -Number(xfrm, "rot") / 60_000.0;
        if (turn != 0) state.Turns.Add(new GroupTurn(frame, turn));
        var childOffset = Child(xfrm, "chOff");
        var childExtent = Child(xfrm, "chExt");
        double cx = Number(childExtent, "cx"), cy = Number(childExtent, "cy");
        foreach (var child in group.Elements())
        {
            if (child.Name.LocalName is not ("pic" or "sp" or "cxnSp" or "grpSp")) continue;
            var childXfrm = child.Name.LocalName switch
            {
                "pic" => Child(Child(child, "spPr"), "xfrm"),
                "grpSp" => Child(Child(child, "grpSpPr"), "xfrm"),
                _ => Child(Child(child, "spPr"), "xfrm"),
            };
            var fraction = cx > 0 && cy > 0 && Child(childXfrm, "off") is { } off && Child(childXfrm, "ext") is { } ext
                ? new FrameFraction((Number(off, "x") - Number(childOffset, "x")) / cx, (Number(off, "y") - Number(childOffset, "y")) / cy, Number(ext, "cx") / cx, Number(ext, "cy") / cy)
                : FrameFraction.Whole;
            ReadObject(state, anchor, frame.Within(fraction), child);
        }
        if (turn != 0) state.Turns.RemoveAt(state.Turns.Count - 1);
    }

    /// <summary>The size and position fields every object takes from its anchor.</summary>
    private static (ImageAnchorKind Kind, AnchorPoint From, AnchorPoint To, double X, double Y, double Width, double Height) Placement(AnchorFrame anchor, XElement? xfrm, FrameFraction frame)
    {
        // A top-level one-cell or absolute anchor without its own size takes it from the object's xfrm; inside groups the anchor decides.
        double width = anchor.HasExtent ? anchor.Width : Number(Child(xfrm, "ext"), "cx") / EmuPerPoint;
        double height = anchor.HasExtent ? anchor.Height : Number(Child(xfrm, "ext"), "cy") / EmuPerPoint;
        return (anchor.Kind, anchor.From, anchor.To, anchor.X, anchor.Y, width, height);
    }

    private static void ReadPicture(State state, AnchorFrame anchor, FrameFraction frame, XElement pic)
    {
        var context = state.Context;
        var sheet = state.Sheet;
        var properties = Child(Child(pic, "nvPicPr"), "cNvPr");
        if ((string?)properties?.Attribute("hidden") is "1" or "true") return;

        var blip = Child(Child(pic, "blipFill"), "blip");
        var embed = blip?.Attributes().FirstOrDefault(a => a.Name.LocalName == "embed")?.Value;
        if (embed is null || !state.Relationships.TryGetValue(embed, out var rel) || rel.IsExternal || !context.Package.Exists(rel.Target))
        {
            context.Options.Warnings.Warn("xlsx.image-missing", $"Sheet '{sheet.Name}' has a picture whose image is missing or linked from outside the file; it was skipped.");
            return;
        }
        ImageData image;
        try { image = context.GetImage(rel.Target); }
        catch (InvalidDataException e)
        {
            context.Options.Warnings.Warn("xlsx.image-unreadable", $"Sheet '{sheet.Name}' has a picture that could not be read: {e.Message}");
            return;
        }

        var src = Child(Child(pic, "blipFill"), "srcRect");
        double Crop(string name) => Math.Clamp(Number(src, name) / 100_000.0, 0, 1);
        var xfrm = Child(Child(pic, "spPr"), "xfrm");
        bool flipH = Flag(xfrm, "flipH"), flipV = Flag(xfrm, "flipV");
        double rotation = -Number(xfrm, "rot") / 60_000.0; // clockwise in the file, counter-clockwise in the model

        var (kind, from, to, x, y, width, height) = Placement(anchor, xfrm, frame);
        sheet.Images.Add(new SheetImage
        {
            Image = image,
            Anchor = kind, From = from, To = to, X = x, Y = y, Width = width, Height = height, Frame = frame,
            CropLeft = Crop("l"), CropTop = Crop("t"), CropRight = Crop("r"), CropBottom = Crop("b"),
            Rotation = state.Turned(rotation), GroupTurns = state.Turns.ToArray(),
            FlipH = flipH, FlipV = flipV,
            ZOrder = state.ZOrder++,
        });
    }

    // ------------------------------------------------------------------ shapes

    private static void ReadShape(State state, AnchorFrame anchor, FrameFraction frame, XElement shape)
    {
        var read = DmlShapes.ReadShape(state, shape, out bool unsupported);
        if (unsupported) state.Shapes++; // a guide formula this reader cannot evaluate
        if (read is null) return;
        var (kind, from, to, x, y, width, height) = Placement(anchor, Child(Child(shape, "spPr"), "xfrm"), frame);
        state.Sheet.Shapes.Add(read with
        {
            Anchor = kind, From = from, To = to, X = x, Y = y, Width = width, Height = height, Frame = frame,
            Rotation = state.Turned(read.Rotation), GroupTurns = state.Turns.ToArray(),
            ZOrder = state.ZOrder++,
        });
    }

    private static AnchorPoint Point(XElement? e) => e is null
        ? default
        : new AnchorPoint(
            (int)Math.Clamp(Number(Child(e, "col"), null), 0, CellReference.MaxColumns - 1), (int)Math.Clamp(Number(Child(e, "row"), null), 0, CellReference.MaxRows - 1),
            Number(Child(e, "colOff"), null) / EmuPerPoint, Number(Child(e, "rowOff"), null) / EmuPerPoint);

    /// <summary>An element's text as a number; 0 when absent or malformed.</summary>
    private static double Number(XElement? e, string? attribute)
    {
        var text = attribute is null ? e?.Value : (string?)e?.Attribute(attribute);
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : 0;
    }

    /// <summary>A chart's cells: references without a sheet name mean the sheet that holds the chart.</summary>
    private sealed class ChartCells(XlsxContext context, Worksheet owner, IReadOnlyList<Worksheet> sheets) : IChartData
    {
        public DmlTheme Theme => context.Theme;

        public bool Date1904 => context.Date1904;

        public bool FramedChartArea => true;

        public Color? TextColor => null;

        public bool PresentationTitle => false;

        public double AutomaticMarkerOutlineWidth => 0.75;

        public bool DarkerMarkerOutline => true;

        public double DefaultLineWidth => 2.25;

        public IEnumerable<Cell?> Cells(string formula)
        {
            foreach (var area in CellReference.ParseAreas(formula))
            {
                if (area.Range is not { } range) continue;
                var sheet = area.Sheet is null ? owner : sheets.FirstOrDefault(s => string.Equals(s.Name, area.Sheet, StringComparison.OrdinalIgnoreCase));
                for (int r = range.FirstRow; r <= range.LastRow; r++)
                    for (int c = range.FirstColumn; c <= range.LastColumn; c++)
                        yield return sheet is not null && sheet.Cells.TryGetValue((r, c), out var cell) ? cell : null;
            }
        }
    }
}
