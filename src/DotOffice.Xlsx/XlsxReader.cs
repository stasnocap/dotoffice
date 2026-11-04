using System.Globalization;
using System.Text;
using System.Xml;
using DotOffice.Core;
using DotOffice.Opc;

namespace DotOffice.Xlsx;

/// <summary>Reads SpreadsheetML (.xlsx) packages into the <see cref="Workbook"/> model.</summary>
public static class XlsxReader
{
    /// <summary>Reads an XLSX package. Malformed or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Workbook Read(Stream stream, XlsxReaderOptions? options = null)
    {
        options ??= new XlsxReaderOptions();
        if (!stream.CanSeek)
        {
            var buffer = new MemoryStream();
            CopyLimited(stream, buffer, options.MaxTotalUncompressedSize);
            buffer.Position = 0;
            stream = buffer;
        }

        try
        {
            using var package = new OpcPackage(stream, new OpcLimits(options.MaxPartSize, options.MaxTotalUncompressedSize, options.MaxPartCount));
            var main = package.FindRelationship("", "officeDocument")
                       ?? throw new InvalidDataException("Package has no main workbook (officeDocument relationship).");
            if (!package.Exists(main.Target))
                throw new InvalidDataException($"Main workbook part '{main.Target}' is missing.");
            return new XlsxContext(package, options, main.Target).ReadWorkbook();
        }
        catch (XmlException e)
        {
            throw new InvalidDataException($"The workbook contains malformed XML: {e.Message}", e);
        }
    }

    private static void CopyLimited(Stream source, Stream target, long limit)
    {
        var buffer = new byte[81920];
        long total = 0;
        int n;
        while ((n = source.Read(buffer)) > 0)
        {
            total += n;
            if (total > limit) throw new InvalidDataException($"Input is larger than the limit of {limit} bytes.");
            target.Write(buffer, 0, n);
        }
    }
}

/// <summary>Small XmlReader helpers shared by the XLSX parsers.</summary>
internal static class XlsxXml
{
    /// <summary>The value of the relationship-id attribute (r:id) whatever namespace prefix is in use.</summary>
    public static string? RelationshipId(this XmlReader r)
    {
        if (!r.HasAttributes) return null;
        for (int i = 0; i < r.AttributeCount; i++)
        {
            r.MoveToAttribute(i);
            bool match = r.LocalName == "id" && r.Prefix.Length > 0;
            var value = r.Value;
            if (match)
            {
                r.MoveToElement();
                return value;
            }
        }
        r.MoveToElement();
        return null;
    }

    public static int? IntAttr(this XmlReader r, string name) =>
        int.TryParse(r.GetAttribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    public static double? DoubleAttr(this XmlReader r, string name) =>
        double.TryParse(r.GetAttribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : null;

    /// <summary>Boolean attribute as written by SpreadsheetML ("1", "true", "0", "false").</summary>
    public static bool BoolAttr(this XmlReader r, string name, bool fallback = false) =>
        r.GetAttribute(name) switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => fallback,
        };

    /// <summary>
    /// Reads a rich-text container (<c>si</c> or <c>is</c>): the text of all <c>t</c> elements in order, except the phonetic
    /// <c>rPh</c> runs, plus the formatting of each <c>r</c> run when there are any. Leaves the reader after the container's end tag.
    /// </summary>
    public static RichString ReadRichText(this XmlReader r)
    {
        if (r.IsEmptyElement)
        {
            r.Read();
            return new RichString("", null);
        }
        int depth = r.Depth;
        var sb = new StringBuilder();
        List<RichRun>? runs = null;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element)
            {
                switch (r.LocalName)
                {
                    case "rPh":
                        r.Skip();
                        continue;
                    case "t":
                        sb.Append(r.ReadElementContentAsString());
                        continue;
                    case "r":
                    {
                        var run = (System.Xml.Linq.XElement)System.Xml.Linq.XNode.ReadFrom(r);
                        var text = new StringBuilder();
                        foreach (var t in run.Elements().Where(e => e.Name.LocalName == "t")) text.Append(t.Value);
                        sb.Append(text);
                        runs ??= [];
                        runs.Add(new RichRun(DecodeEscapes(text.ToString()), run.Elements().FirstOrDefault(e => e.Name.LocalName == "rPr")));
                        continue;
                    }
                }
            }
            r.Read();
        }
        r.Read(); // past the end tag
        return new RichString(DecodeEscapes(sb.ToString()), runs);
    }

    /// <summary>Decodes <c>_xHHHH_</c> escapes that Excel uses for control characters in strings.</summary>
    public static string DecodeEscapes(string text)
    {
        if (!text.Contains("_x", StringComparison.Ordinal)) return text;
        var sb = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == '_' && i + 6 < text.Length && text[i + 1] == 'x' && text[i + 6] == '_' &&
                ushort.TryParse(text.AsSpan(i + 2, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var code))
            {
                sb.Append((char)code);
                i += 6;
            }
            else
            {
                sb.Append(text[i]);
            }
        }
        return sb.ToString();
    }
}

/// <summary>A formatted run of a rich string; <see cref="Properties"/> is the raw <c>rPr</c> element, if any.</summary>
internal sealed record RichRun(string Text, System.Xml.Linq.XElement? Properties);

/// <summary>A shared or inline string: its plain text and, when it has formatted runs, the runs.</summary>
internal sealed record RichString(string Text, List<RichRun>? Runs);
