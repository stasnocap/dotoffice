using System.Text;
using System.Xml.Linq;
using DotOffice.Opc;

namespace DotOffice.Odf;

/// <summary>Opening an OpenDocument package: limits, encryption and type checks, and the styles part.</summary>
internal static class OdfPackage
{
    public static OpcPackage Open(Stream stream, OdfReaderOptions options, string mimePrefix, out OdfStyles styles)
    {
        if (!stream.CanSeek)
        {
            var buffer = new MemoryStream();
            var chunk = new byte[81920];
            long total = 0;
            int n;
            while ((n = stream.Read(chunk)) > 0)
            {
                total += n;
                if (total > options.MaxTotalUncompressedSize) throw new InvalidDataException("Input is larger than the limit.");
                buffer.Write(chunk, 0, n);
            }
            buffer.Position = 0;
            stream = buffer;
        }
        var package = new OpcPackage(stream, new OpcLimits(options.MaxPartSize, options.MaxTotalUncompressedSize, options.MaxPartCount));
        try
        {
            if (!package.Exists("content.xml")) throw new InvalidDataException("Package has no content.xml; it is not an OpenDocument file.");
            if (package.Exists("META-INF/manifest.xml"))
            {
                using var r = package.OpenXml("META-INF/manifest.xml");
                if (XElement.Load(r).Descendants().Any(e => e.Name.LocalName == "encryption-data"))
                    throw new InvalidDataException("The document is encrypted.");
            }
            if (package.Exists("mimetype"))
            {
                var mime = Encoding.ASCII.GetString(package.ReadAllBytes("mimetype")).Trim();
                if (mime.Length > 0 && !mime.StartsWith(mimePrefix, StringComparison.Ordinal))
                    throw new InvalidDataException($"The package is '{mime}', not '{mimePrefix}'.");
            }
            styles = new OdfStyles();
            if (package.Exists("styles.xml"))
            {
                using var r = package.OpenXml("styles.xml");
                foreach (var part in XElement.Load(r).Elements().Where(e => e.Name.LocalName is "styles" or "automatic-styles" or "master-styles" or "font-face-decls"))
                    styles.Add(part);
            }
            return package;
        }
        catch
        {
            package.Dispose();
            throw;
        }
    }
}
