using System.Globalization;
using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Namespaces and streaming XmlReader helpers shared by all part readers.</summary>
internal static class WordXml
{
    public const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public const string WStrict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    public const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string RStrict = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    public const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string AStrict = "http://purl.oclc.org/ooxml/drawingml/main";
    public const string WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public const string WPStrict = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
    public const string MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    public static XmlReaderSettings Settings(long maxCharacters) => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        MaxCharactersInDocument = maxCharacters,
        MaxCharactersFromEntities = 1024,
        CloseInput = true,
    };

    public static bool IsW(this XmlReader r) => r.NamespaceURI is W or WStrict;

    public static bool IsA(this XmlReader r) => r.NamespaceURI is A or AStrict;

    public static bool IsWP(this XmlReader r) => r.NamespaceURI is WP or WPStrict;

    public static bool IsMC(this XmlReader r) => r.NamespaceURI == MC;

    /// <summary>Moves to the first child node. Returns false (having consumed the element) when it is empty.</summary>
    public static bool Enter(this XmlReader r)
    {
        if (r.IsEmptyElement)
        {
            r.Read();
            return false;
        }
        r.Read();
        return true;
    }

    /// <summary>
    /// Advances to the next child element of the element at <paramref name="depth"/>; returns false after
    /// consuming the parent's end tag. Each handled child must be fully consumed (e.g. via Skip()).
    /// </summary>
    public static bool NextChild(this XmlReader r, int depth)
    {
        while (r.Depth > depth && !r.EOF)
        {
            if (r.NodeType == XmlNodeType.Element) return true;
            r.Read();
        }
        if (r.NodeType == XmlNodeType.EndElement && r.Depth == depth) r.Read();
        return false;
    }

    /// <summary>Gets a WordprocessingML attribute (w:name), tolerating strict namespaces and unprefixed names.</summary>
    public static string? WAttr(this XmlReader r, string localName) =>
        r.GetAttribute(localName, W) ?? r.GetAttribute(localName, WStrict) ?? r.GetAttribute(localName);

    public static string? RAttr(this XmlReader r, string localName) =>
        r.GetAttribute(localName, R) ?? r.GetAttribute(localName, RStrict);

    /// <summary>ST_OnOff: absent val means true.</summary>
    public static bool OnOff(this XmlReader r, string attribute = "val") =>
        r.WAttr(attribute) is not { } v || v is "1" or "true" or "on";

    public static int? IntAttr(this XmlReader r, string localName) =>
        int.TryParse(r.WAttr(localName), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    public static double? DoubleAttr(this XmlReader r, string localName) =>
        double.TryParse(r.WAttr(localName), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : null;

    /// <summary>Parses ST_TwipsMeasure / ST_SignedTwipsMeasure (twips or a universal measure) into points.</summary>
    public static double? TwipsAttr(this XmlReader r, string localName) => ParseTwips(r.WAttr(localName));

    public static double? ParseTwips(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        value = value.Trim();
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips))
            return double.IsFinite(twips) ? Units.FromTwips(twips) : null;
        if (value.Length < 3) return null;
        var unit = value[^2..];
        if (!double.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var n) || !double.IsFinite(n))
            return null;
        return unit switch
        {
            "pt" => n,
            "in" => Units.FromInches(n),
            "mm" => Units.FromMillimeters(n),
            "cm" => Units.FromMillimeters(n * 10),
            "pc" or "pi" => n * 12,
            _ => null,
        };
    }

    /// <summary>Reads all text of the current element and moves past it.</summary>
    public static string ReadText(this XmlReader r) => r.IsEmptyElement ? Consume(r, "") : r.ReadElementContentAsString();

    private static string Consume(XmlReader r, string value)
    {
        r.Read();
        return value;
    }
}