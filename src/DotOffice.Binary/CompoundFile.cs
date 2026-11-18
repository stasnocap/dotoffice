using System.Buffers.Binary;
using System.Text;

namespace DotOffice.Binary;

/// <summary>Limits applied while reading binary Office files.</summary>
public sealed class BinaryReaderOptions
{
    /// <summary>Receives warnings about unsupported or malformed content.</summary>
    public DotOffice.Core.IConversionWarningSink? Warnings { get; init; }

    /// <summary>Maximum size of the whole file.</summary>
    public long MaxFileSize { get; init; } = 512L * 1024 * 1024;

    /// <summary>Maximum size of any one stream read out of the file.</summary>
    public long MaxStreamSize { get; init; } = 256L * 1024 * 1024;

    /// <summary>Maximum number of non-empty cells read from a workbook.</summary>
    public long MaxCellCount { get; init; } = 5_000_000;
}

/// <summary>
/// A Compound File Binary (OLE2) container, the wrapper of .xls, .doc and .ppt files. The file is read into memory (bounded by
/// <see cref="BinaryReaderOptions.MaxFileSize"/>); sector chains are followed with a step limit, so a looping or oversized chain throws
/// <see cref="InvalidDataException"/> instead of hanging or reading without bound.
/// </summary>
public sealed class CompoundFile
{
    private const uint EndOfChain = 0xFFFFFFFE, FreeSector = 0xFFFFFFFF, MaxRegular = 0xFFFFFFFA;
    private static readonly byte[] Signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private readonly byte[] _data;
    private readonly int _sectorSize, _miniSectorSize, _sectorCount;
    private readonly long _miniCutoff, _maxStream;
    private readonly uint[] _fat;
    private uint[] _miniFat = [];
    private byte[]? _miniStream;
    private readonly List<Entry> _entries = [];

    private sealed record Entry(string Name, int Type, uint Left, uint Right, uint Child, uint Start, long Size);

    /// <summary>Whether <paramref name="header"/> (at least 8 bytes) starts like a compound file.</summary>
    public static bool HasSignature(ReadOnlySpan<byte> header) => header.Length >= 8 && header[..8].SequenceEqual(Signature);

    public CompoundFile(Stream stream, BinaryReaderOptions? options = null)
    {
        options ??= new BinaryReaderOptions();
        _maxStream = options.MaxStreamSize;
        if (stream.CanSeek && stream.Length - stream.Position > options.MaxFileSize) throw new InvalidDataException("The file is larger than the limit.");
        var buffer = new MemoryStream();
        var chunk = new byte[81920];
        int n;
        while ((n = stream.Read(chunk)) > 0)
        {
            if (buffer.Length + n > options.MaxFileSize) throw new InvalidDataException("The file is larger than the limit.");
            buffer.Write(chunk, 0, n);
        }
        _data = buffer.ToArray();
        if (_data.Length < 512 || !HasSignature(_data)) throw new InvalidDataException("The file is not a compound (OLE2) file.");

        int sectorShift = BinaryPrimitives.ReadUInt16LittleEndian(_data.AsSpan(30)), miniShift = BinaryPrimitives.ReadUInt16LittleEndian(_data.AsSpan(32));
        if (sectorShift is not (9 or 12) || miniShift is < 3 or > 9) throw new InvalidDataException("The compound file has an unsupported sector size.");
        _sectorSize = 1 << sectorShift;
        _miniSectorSize = 1 << miniShift;
        _sectorCount = (int)Math.Min(int.MaxValue, Math.Max(0, (_data.Length - _sectorSize) / _sectorSize + 1));
        uint fatCount = U32(44), firstDirectory = U32(48), firstMiniFat = U32(60), miniFatCount = U32(64), firstDifat = U32(68), difatCount = U32(72);
        _miniCutoff = U32(56);
        if (fatCount > _sectorCount || miniFatCount > _sectorCount || difatCount > _sectorCount) throw new InvalidDataException("The compound file's allocation tables are larger than the file.");

        // The FAT sectors are listed in the header (109 of them) and then in a chain of DIFAT sectors.
        var fatSectors = new List<uint>();
        for (int i = 0; i < 109 && fatSectors.Count < fatCount; i++)
            if (U32(76 + 4 * i) is var s && s < MaxRegular) fatSectors.Add(s);
        uint difat = firstDifat;
        for (uint steps = 0; difat < MaxRegular && fatSectors.Count < fatCount; steps++)
        {
            if (steps > difatCount) throw new InvalidDataException("The compound file's DIFAT chain loops.");
            var span = Sector(difat);
            for (int i = 0; i < _sectorSize / 4 - 1 && fatSectors.Count < fatCount; i++)
                if (BinaryPrimitives.ReadUInt32LittleEndian(span[(4 * i)..]) is var s && s < MaxRegular) fatSectors.Add(s);
            difat = BinaryPrimitives.ReadUInt32LittleEndian(span[(_sectorSize - 4)..]);
        }
        _fat = new uint[fatSectors.Count * (_sectorSize / 4)];
        for (int i = 0; i < fatSectors.Count; i++)
        {
            var span = Sector(fatSectors[i]);
            for (int k = 0; k < _sectorSize / 4; k++) _fat[i * (_sectorSize / 4) + k] = BinaryPrimitives.ReadUInt32LittleEndian(span[(4 * k)..]);
        }

        var directory = ReadChain(firstDirectory, long.MaxValue, regularOnly: true);
        for (int at = 0; at + 128 <= directory.Length && _entries.Count < 100_000; at += 128)
        {
            var e = directory.AsSpan(at, 128);
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(e[64..]);
            string name = nameLength is >= 2 and <= 64 ? Encoding.Unicode.GetString(e[..(nameLength - 2)]) : "";
            long size = BinaryPrimitives.ReadInt64LittleEndian(e[120..]);
            if (_sectorSize == 512) size &= 0xFFFFFFFF;
            _entries.Add(new Entry(name, e[66], BinaryPrimitives.ReadUInt32LittleEndian(e[68..]), BinaryPrimitives.ReadUInt32LittleEndian(e[72..]),
                BinaryPrimitives.ReadUInt32LittleEndian(e[76..]), BinaryPrimitives.ReadUInt32LittleEndian(e[116..]), size));
        }
        if (_entries.Count == 0 || _entries[0].Type != 5) throw new InvalidDataException("The compound file has no root entry.");

        if (miniFatCount > 0 && firstMiniFat < MaxRegular)
        {
            var bytes = ReadChain(firstMiniFat, (long)miniFatCount * _sectorSize + _sectorSize, regularOnly: true);
            _miniFat = new uint[bytes.Length / 4];
            for (int i = 0; i < _miniFat.Length; i++) _miniFat[i] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(4 * i));
        }
        // The mini stream is the root entry's own stream, which holds the small streams.
        if (_entries[0].Start < MaxRegular && _entries[0].Size > 0)
            _miniStream = ReadChain(_entries[0].Start, Math.Min(_entries[0].Size, options.MaxFileSize), regularOnly: true, exactSize: Math.Min(_entries[0].Size, options.MaxFileSize));
    }

    private uint U32(int at) => BinaryPrimitives.ReadUInt32LittleEndian(_data.AsSpan(at));

    private ReadOnlySpan<byte> Sector(uint index)
    {
        if (index >= _sectorCount) throw new InvalidDataException("A sector number of the compound file is outside the file.");
        long at = (index + 1L) * _sectorSize;
        if (at + _sectorSize > _data.Length)
        {
            // The last sector may be short in files written by some tools.
            if (at >= _data.Length) throw new InvalidDataException("A sector of the compound file is outside the file.");
            var padded = new byte[_sectorSize];
            Array.Copy(_data, at, padded, 0, _data.Length - at);
            return padded;
        }
        return _data.AsSpan((int)at, _sectorSize);
    }

    private byte[] ReadChain(uint start, long limit, bool regularOnly, long exactSize = -1)
    {
        var result = new MemoryStream();
        uint sector = start;
        long steps = 0;
        while (sector < MaxRegular)
        {
            if (++steps > _sectorCount) throw new InvalidDataException("A sector chain of the compound file loops.");
            if (result.Length >= limit && limit != long.MaxValue) throw new InvalidDataException("A stream of the compound file is larger than the limit.");
            result.Write(Sector(sector));
            sector = sector < _fat.Length ? _fat[sector] : EndOfChain;
        }
        _ = regularOnly;
        var bytes = result.ToArray();
        if (exactSize >= 0)
        {
            if (bytes.Length < exactSize) throw new InvalidDataException("A stream of the compound file is shorter than its length says.");
            Array.Resize(ref bytes, (int)exactSize);
        }
        return bytes;
    }

    private byte[] ReadMiniChain(uint start, long size)
    {
        if (_miniStream is null) throw new InvalidDataException("The compound file has small streams but no mini stream.");
        var result = new byte[size];
        long written = 0;
        uint sector = start;
        long steps = 0;
        while (written < size)
        {
            if (sector >= MaxRegular || ++steps > _miniFat.Length) throw new InvalidDataException("A small stream of the compound file is truncated or loops.");
            long at = (long)sector * _miniSectorSize;
            if (at >= _miniStream.Length) throw new InvalidDataException("A small stream of the compound file is outside the mini stream.");
            int n = (int)Math.Min(Math.Min(_miniSectorSize, size - written), _miniStream.Length - at);
            Array.Copy(_miniStream, at, result, written, n);
            written += n;
            sector = sector < _miniFat.Length ? _miniFat[sector] : EndOfChain;
        }
        return result;
    }

    /// <summary>The names of the streams and storages directly in the root, for diagnostics and detection.</summary>
    public IEnumerable<string> RootNames => Children(0).Select(i => _entries[i].Name);

    private List<int> Children(int parent)
    {
        var result = new List<int>();
        var seen = new HashSet<uint>();
        var stack = new Stack<uint>();
        if (_entries[parent].Child < _entries.Count) stack.Push(_entries[parent].Child);
        while (stack.Count > 0)
        {
            uint id = stack.Pop();
            if (id >= _entries.Count || !seen.Add(id)) continue;
            result.Add((int)id);
            stack.Push(_entries[(int)id].Left);
            stack.Push(_entries[(int)id].Right);
        }
        return result;
    }

    private int Find(int parent, string name)
    {
        foreach (var i in Children(parent))
            if (string.Equals(_entries[i].Name, name, StringComparison.OrdinalIgnoreCase)) return i;
        return -1;
    }

    /// <summary>Whether the root has a stream or storage of the name.</summary>
    public bool Contains(string name) => Find(0, name) >= 0;

    /// <summary>The bytes of the root-level stream <paramref name="name"/>, or null when there is none.</summary>
    public byte[]? ReadStream(string name)
    {
        int i = Find(0, name);
        if (i < 0 || _entries[i].Type != 2) return null;
        var e = _entries[i];
        if (e.Size > _maxStream) throw new InvalidDataException($"The stream '{name}' is larger than the limit.");
        if (e.Size == 0) return [];
        return e.Size < _miniCutoff ? ReadMiniChain(e.Start, e.Size) : ReadChain(e.Start, e.Size, regularOnly: false, exactSize: e.Size);
    }
}
