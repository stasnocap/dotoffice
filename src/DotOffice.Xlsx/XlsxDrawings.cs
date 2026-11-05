using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>Reads the pictures of a worksheet's drawing part (SpreadsheetML Drawing, <c>xdr:wsDr</c>).</summary>
internal static class XlsxDrawings
{
    private const double EmuPerPoint = 12700;

    /// <summary>What a drawing contained that is not rendered.</summary>
    internal sealed record Skipped(int Charts, int Shapes, int Other)
    {
        public bool Any => Charts + Shapes + Other > 0;
    }

    /// <summary>
    /// Adds the pictures of the drawing part to <paramref name="sheet"/> and reports what was left out. Pictures that are hidden or
    /// excluded from printing (<c>fPrintsWithSheet="0"</c>, as Excel's "Print object" option writes it) are not printed.
    /// </summary>
    public static Skipped Read(XlsxContext context, Worksheet sheet, string drawingPart)
    {
        int charts = 0, shapes = 0, other = 0;
        XDocument doc;
        using (var reader = context.Package.OpenXml(drawingPart)) doc = XDocument.Load(reader);
        if (doc.Root is not { } root || root.Name.LocalName != "wsDr") return new Skipped(0, 0, 1);
        var rels = context.Package.GetRelationships(drawingPart);

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

            foreach (var content in anchor.Elements())
            {
                switch (content.Name.LocalName)
                {
                    case "pic":
                        ReadPicture(context, sheet, anchor, content, kind.Value, rels);
                        break;
                    case "graphicFrame": charts++; break;
                    case "sp" or "cxnSp" or "grpSp": shapes++; break;
                    case "from" or "to" or "pos" or "ext" or "clientData": break;
                    default: other++; break; // AlternateContent (slicers, …), contentPart
                }
            }
        }
        return new Skipped(charts, shapes, other);
    }

    private static void ReadPicture(XlsxContext context, Worksheet sheet, XElement anchor, XElement pic, ImageAnchorKind kind,
        IReadOnlyDictionary<string, Opc.Relationship> rels)
    {
        var properties = Child(Child(pic, "nvPicPr"), "cNvPr");
        if ((string?)properties?.Attribute("hidden") is "1" or "true") return;
        if ((string?)Child(anchor, "clientData")?.Attribute("fPrintsWithSheet") is "0" or "false") return;

        var blip = Child(Child(pic, "blipFill"), "blip");
        var embed = blip?.Attributes().FirstOrDefault(a => a.Name.LocalName == "embed")?.Value;
        if (embed is null || !rels.TryGetValue(embed, out var rel) || rel.IsExternal || !context.Package.Exists(rel.Target))
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
        bool flipH = (string?)xfrm?.Attribute("flipH") is "1" or "true", flipV = (string?)xfrm?.Attribute("flipV") is "1" or "true";
        double rotation = -Number(xfrm, "rot") / 60_000.0; // clockwise in the file, counter-clockwise in the model

        var ext = Child(anchor, "ext") ?? Child(xfrm, "ext");
        double width = Number(ext, "cx") / EmuPerPoint, height = Number(ext, "cy") / EmuPerPoint;
        var pos = Child(anchor, "pos");
        sheet.Images.Add(new SheetImage
        {
            Image = image,
            Anchor = kind,
            From = Point(Child(anchor, "from")),
            To = Point(Child(anchor, "to")),
            X = Number(pos, "x") / EmuPerPoint,
            Y = Number(pos, "y") / EmuPerPoint,
            Width = width,
            Height = height,
            CropLeft = Crop("l"), CropTop = Crop("t"), CropRight = Crop("r"), CropBottom = Crop("b"),
            Rotation = rotation,
            FlipH = flipH, FlipV = flipV,
        });
    }

    private static AnchorPoint Point(XElement? e) => e is null
        ? default
        : new AnchorPoint(
            (int)Math.Clamp(Number(Child(e, "col")), 0, CellReference.MaxColumns - 1), (int)Math.Clamp(Number(Child(e, "row")), 0, CellReference.MaxRows - 1),
            Number(Child(e, "colOff")) / EmuPerPoint, Number(Child(e, "rowOff")) / EmuPerPoint);

    private static XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(e => e.Name.LocalName == name);

    /// <summary>An element's text or an attribute's value as a number; 0 when absent or malformed.</summary>
    private static double Number(XElement? e, string? attribute)
    {
        var text = attribute is null ? e?.Value : (string?)e?.Attribute(attribute);
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : 0;
    }

    private static double Number(XElement? e) => Number(e, null);
}
