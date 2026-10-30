using System.IO.Compression;
using System.Xml;

namespace DotOffice.Docx;

internal sealed record Relationship(string Id, string Type, string Target, bool IsExternal)
{
    /// <summary>Last segment of the relationship type URI, e.g. "styles" or "officeDocument".</summary>
    public string Kind => Type[(Type.LastIndexOf('/') + 1)..];
}

/// <summary>
/// Read-only Open Packaging Conventions container over a zip archive. Every read is metered against
/// the configured limits, so lying zip headers cannot inflate memory use (zip bomb protection).
/// </summary>
internal sealed class OpcPackage : IDisposable
{
    private readonly ZipArchive _zip;
    private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, Relationship>> _relationships = new(StringComparer.OrdinalIgnoreCase);
    private readonly DocxReaderOptions _options;
    private long _bytesRead;

    public OpcPackage(Stream stream, DocxReaderOptions options)
    {
        _options = options;
        try
        {
            _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        }
        catch (InvalidDataException e)
        {
            throw new InvalidDataException("The input is not a valid DOCX (zip) package.", e);
        }
        if (_zip.Entries.Count > options.MaxPartCount)
            throw new InvalidDataException($"Package has {_zip.Entries.Count} parts, more than the limit of {options.MaxPartCount}.");
        foreach (var entry in _zip.Entries)
            _entries[Normalize(entry.FullName)] = entry;
    }

    public bool Exists(string partName) => _entries.ContainsKey(Normalize(partName));

    /// <summary>Opens a part for reading; the stream throws once the per-part or total limit is exceeded.</summary>
    public Stream Open(string partName)
    {
        if (!_entries.TryGetValue(Normalize(partName), out var entry))
            throw new FileNotFoundException($"Package part '{partName}' does not exist.");
        if (entry.Length > _options.MaxPartSize)
            throw new InvalidDataException($"Part '{partName}' is larger than the limit of {_options.MaxPartSize} bytes.");
        return new MeteredStream(entry.Open(), _options.MaxPartSize, this, partName);
    }

    public byte[] ReadAllBytes(string partName)
    {
        using var stream = Open(partName);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    public XmlReader OpenXml(string partName) =>
        XmlReader.Create(Open(partName), WordXml.Settings(_options.MaxPartSize));

    /// <summary>Relationships of a part ("" for package-level relationships), keyed by id.</summary>
    public IReadOnlyDictionary<string, Relationship> GetRelationships(string partName)
    {
        partName = Normalize(partName);
        if (_relationships.TryGetValue(partName, out var cached)) return cached;

        var result = new Dictionary<string, Relationship>(StringComparer.Ordinal);
        int slash = partName.LastIndexOf('/');
        var relsPart = slash < 0 ? $"_rels/{partName}.rels" : $"{partName[..slash]}/_rels/{partName[(slash + 1)..]}.rels";
        if (partName.Length == 0) relsPart = "_rels/.rels";
        if (Exists(relsPart))
        {
            using var r = OpenXml(relsPart);
            while (r.Read())
            {
                if (r.NodeType != XmlNodeType.Element || r.LocalName != "Relationship") continue;
                var id = r.GetAttribute("Id");
                var type = r.GetAttribute("Type");
                var target = r.GetAttribute("Target");
                if (id is null || type is null || target is null) continue;
                bool external = string.Equals(r.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                result[id] = new Relationship(id, type, external ? target : Resolve(partName, target), external);
            }
        }
        _relationships[partName] = result;
        return result;
    }

    public Relationship? FindRelationship(string partName, string kind) =>
        GetRelationships(partName).Values.FirstOrDefault(r => !r.IsExternal && r.Kind == kind);

    public void Dispose() => _zip.Dispose();

    private void Account(long bytes, string partName)
    {
        _bytesRead += bytes;
        if (_bytesRead > _options.MaxTotalUncompressedSize)
            throw new InvalidDataException(
                $"Package expands beyond the limit of {_options.MaxTotalUncompressedSize} bytes (while reading '{partName}').");
    }

    internal static string Normalize(string partName)
    {
        var name = Uri.UnescapeDataString(partName.Replace('\\', '/')).TrimStart('/');
        return name;
    }

    /// <summary>Resolves a relationship target relative to the source part's folder.</summary>
    internal static string Resolve(string sourcePart, string target)
    {
        if (target.StartsWith('/')) return Normalize(target);
        int slash = sourcePart.LastIndexOf('/');
        var segments = new List<string>(slash < 0 ? [] : sourcePart[..slash].Split('/'));
        foreach (var segment in target.Replace('\\', '/').Split('/'))
        {
            if (segment is "" or ".") continue;
            if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
            else segments.Add(segment);
        }
        return Normalize(string.Join('/', segments));
    }

    private sealed class MeteredStream(Stream inner, long limit, OpcPackage package, string partName) : Stream
    {
        private long _read;

        public override int Read(byte[] buffer, int offset, int count) => Count(inner.Read(buffer, offset, count));

        public override int Read(Span<byte> buffer) => Count(inner.Read(buffer));

        private int Count(int n)
        {
            _read += n;
            if (_read > limit)
                throw new InvalidDataException($"Part '{partName}' expands beyond the limit of {limit} bytes.");
            package.Account(n, partName);
            return n;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => _read; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing) inner.Dispose();
            base.Dispose(disposing);
        }
    }
}