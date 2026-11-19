using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Reads the font table part (w:fonts): each font's alternative name and whether it is of the roman family.</summary>
internal static class FontTable
{
    public static Dictionary<string, FontDeclaration> Read(XmlReader r)
    {
        var fonts = new Dictionary<string, FontDeclaration>(StringComparer.OrdinalIgnoreCase);
        string? name = null, altName = null;
        bool roman = false;
        int depth = -1;
        void Flush()
        {
            if (!string.IsNullOrEmpty(name)) fonts.TryAdd(name, new FontDeclaration(string.IsNullOrEmpty(altName) ? null : altName, roman));
            name = altName = null; roman = false;
        }
        while (r.Read())
        {
            if (r.NodeType != XmlNodeType.Element || !r.IsW()) continue;
            if (r.LocalName == "font")
            {
                Flush();
                name = r.WAttr("name");
                depth = r.Depth;
            }
            else if (name is not null && r.Depth == depth + 1)
            {
                if (r.LocalName == "altName") altName = r.WAttr("val");
                else if (r.LocalName == "family") roman = r.WAttr("val") == "roman";
            }
        }
        Flush();
        return fonts;
    }
}
