using System.Buffers.Binary;
using System.Text;
using DotOffice.Binary;
using DotOffice.Core;

namespace DotOffice.Tests.Binary;

public class XlsReaderTests
{
    private static byte[] Rec(int id, params byte[] data)
    {
        var bytes = new byte[4 + data.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(bytes, (ushort)id);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(2), (ushort)data.Length);
        data.CopyTo(bytes, 4);
        return bytes;
    }

    private static byte[] U16(params int[] values) => values.SelectMany(v => BitConverter.GetBytes((ushort)v)).ToArray();

    private static byte[] Concat(params byte[][] parts) => parts.SelectMany(p => p).ToArray();

    /// <summary>A compound file holding one stream, laid out as header, FAT sector, directory sector and the stream's sectors.</summary>
    internal static byte[] Cfb(string name, byte[] stream, bool loop = false)
    {
        int sectors = (stream.Length + 511) / 512;
        var file = new byte[512 * (3 + sectors)];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(file, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(24), 0x3E);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(26), 3);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(28), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(30), 9);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(32), 6);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(44), 1);   // FAT sectors
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(48), 1);   // first directory sector
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(56), 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(60), 0xFFFFFFFE);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(68), 0xFFFFFFFE);
        for (int i = 0; i < 109; i++) BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(76 + 4 * i), i == 0 ? 0u : 0xFFFFFFFF);

        var fat = file.AsSpan(512, 512);
        for (int i = 0; i < 128; i++) BinaryPrimitives.WriteUInt32LittleEndian(fat[(4 * i)..], 0xFFFFFFFF);
        BinaryPrimitives.WriteUInt32LittleEndian(fat, 0xFFFFFFFD);
        BinaryPrimitives.WriteUInt32LittleEndian(fat[4..], 0xFFFFFFFE);
        for (int i = 0; i < sectors; i++) BinaryPrimitives.WriteUInt32LittleEndian(fat[(4 * (2 + i))..], i == sectors - 1 ? (loop ? 2u : 0xFFFFFFFE) : (uint)(3 + i));

        void Entry(int index, string entryName, int type, uint child, uint start, long size)
        {
            var e = file.AsSpan(1024 + 128 * index, 128);
            var nameBytes = Encoding.Unicode.GetBytes(entryName);
            nameBytes.CopyTo(e);
            BinaryPrimitives.WriteUInt16LittleEndian(e[64..], (ushort)(nameBytes.Length + 2));
            e[66] = (byte)type;
            BinaryPrimitives.WriteUInt32LittleEndian(e[68..], 0xFFFFFFFF);
            BinaryPrimitives.WriteUInt32LittleEndian(e[72..], 0xFFFFFFFF);
            BinaryPrimitives.WriteUInt32LittleEndian(e[76..], child);
            BinaryPrimitives.WriteUInt32LittleEndian(e[116..], start);
            BinaryPrimitives.WriteInt64LittleEndian(e[120..], size);
        }
        Entry(0, "Root Entry", 5, 1, 0xFFFFFFFE, 0);
        Entry(1, name, 2, 0xFFFFFFFF, 2, stream.Length);
        stream.CopyTo(file, 512 * 3);
        return file;
    }

    private static byte[] Workbook(bool encrypted = false)
    {
        var bof = (int type) => Rec(0x0809, U16(0x0600, type, 0, 0, 0, 0, 0, 0));
        var sst = Rec(0x00FC, Concat(BitConverter.GetBytes(1), BitConverter.GetBytes(1), U16(5), [0], Encoding.Latin1.GetBytes("Hello")));
        byte[] Globals(int sheetOffset) => Concat(bof(5), encrypted ? Rec(0x002F, U16(0, 0, 0)) : [], sst,
            Rec(0x0085, Concat(BitConverter.GetBytes(sheetOffset), [0, 0, 4, 0], Encoding.Latin1.GetBytes("Data"))), Rec(0x000A));
        int offset = Globals(0).Length;
        var sheet = Concat(bof(0x10),
            Rec(0x00FD, Concat(U16(0, 0, 0), BitConverter.GetBytes(0))),
            Rec(0x0203, Concat(U16(0, 1, 0), BitConverter.GetBytes(3.5))),
            Rec(0x027E, Concat(U16(1, 0, 0), BitConverter.GetBytes((12 << 2) | 2 | 1))), // RK: 12 / 100
            Rec(0x00E5, Concat(U16(1), U16(2, 3, 0, 1))),
            Rec(0x000A));
        var all = Concat(Globals(offset), sheet);
        // Compound files keep streams under 4096 bytes in the mini stream, which this helper does not build.
        return all.Length >= 4096 ? all : Concat(all, new byte[4096 - all.Length]);
    }

    [Fact]
    public void ReadsCellsMergesAndSheetNames()
    {
        var wb = XlsReader.Read(new MemoryStream(Cfb("Workbook", Workbook())));
        var sheet = Assert.Single(wb.Sheets);
        Assert.Equal("Data", sheet.Name);
        Assert.Equal("Hello", sheet.Cells[(0, 0)].Text);
        Assert.Equal("3.5", sheet.Cells[(0, 1)].Text);
        Assert.Equal(0.12, sheet.Cells[(1, 0)].Value);
        Assert.Equal(new CellRange(2, 0, 3, 1), Assert.Single(sheet.Merges));
    }

    [Fact]
    public void RejectsEncryptedWorkbooks() =>
        Assert.Throws<InvalidDataException>(() => XlsReader.Read(new MemoryStream(Cfb("Workbook", Workbook(encrypted: true)))));

    [Fact]
    public void RejectsALoopingSectorChain() =>
        Assert.Throws<InvalidDataException>(() => XlsReader.Read(new MemoryStream(Cfb("Workbook", Workbook(), loop: true))));

    [Fact]
    public void RejectsFilesThatAreNotCompoundFiles() =>
        Assert.Throws<InvalidDataException>(() => XlsReader.Read(new MemoryStream(new byte[2048])));

    [Fact]
    public void RejectsStreamsLargerThanTheLimit() =>
        Assert.Throws<InvalidDataException>(() => XlsReader.Read(new MemoryStream(Cfb("Workbook", Workbook())), new BinaryReaderOptions { MaxStreamSize = 1000 }));
}
