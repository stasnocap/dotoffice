using System.Text;
using System.Xml.Linq;
using DotOffice.Opc;

namespace DotOffice.Odf;

/// <summary>Opening an OpenDocument package: limits, encryption and type checks, and the styles part.</summary>
internal static class OdfPackage
{
    /// <summary>
    /// A flat OpenDocument file (<c>.fods</c>, <c>.fodt</c>, <c>.fodp</c>: one XML document holding styles, content and pictures) is wrapped into
    /// an in-memory package whose content and styles parts are that document, so the readers treat it like a zip.
    /// </summary>
    private static Stream FlatToPackage(Stream stream, OdfReaderOptions options)
    {
        var start = stream.Position;
        int b = stream.ReadByte();
        if (b == 0xEF) { stream.Position = start + 3; b = stream.ReadByte(); }
        while (b is ' ' or '\t' or '\r' or '\n') b = stream.ReadByte();
        stream.Position = start;
        if (b != '<') return stream;
        var bytes = new byte[stream.Length - start];
        stream.ReadExactly(bytes);
        XElement root;
        try { root = XElement.Load(new MemoryStream(bytes)); }
        catch (System.Xml.XmlException e) { throw new InvalidDataException($"The document contains malformed XML: {e.Message}", e); }
        if (root.Name != Ns.Office + "document") throw new InvalidDataException("The XML file is not a flat OpenDocument file.");
        var ms = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, byte[] data)
            {
                using var w = zip.CreateEntry(name).Open();
                w.Write(data);
            }
            if ((string?)root.Attribute(Ns.Office + "mimetype") is { } mime) Add("mimetype", Encoding.ASCII.GetBytes(mime));
            Add("content.xml", bytes);
            Add("styles.xml", bytes);
        }
        ms.Position = 0;
        return ms;
    }

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
        stream = FlatToPackage(stream, options);
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
