using System.Xml;
using DotOffice.Core;
using DotOffice.Opc;

namespace DotOffice.Pptx;

/// <summary>Reads PresentationML (.pptx) packages into the <see cref="Presentation"/> model.</summary>
public static class PptxReader
{
    /// <summary>Reads a PPTX package. Malformed or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Presentation Read(Stream stream, PptxReaderOptions? options = null)
    {
        options ??= new PptxReaderOptions();
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
                       ?? throw new InvalidDataException("Package has no main presentation (officeDocument relationship).");
            if (!package.Exists(main.Target))
                throw new InvalidDataException($"Main presentation part '{main.Target}' is missing.");
            return new PptxContext(package, options).Read(main.Target);
        }
        catch (XmlException e)
        {
            throw new InvalidDataException($"The presentation contains malformed XML: {e.Message}", e);
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
