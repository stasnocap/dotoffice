using System.Buffers.Binary;
using System.Text;
using DotOffice.Binary;
using DotOffice.Core;

namespace DotOffice.Tests.Binary;

public class DocReaderTests
{
    /// <summary>A compound file with up to three streams: header, FAT sector, directory sector, then each stream's sectors in turn.</summary>
    internal static byte[] Cfb(params (string Name, byte[] Data)[] streams)
    {
        // Streams under 4096 bytes would belong in the mini stream, which this helper does not build.
        streams = streams.Select(s => (s.Name, s.Data.Length >= 4096 ? s.Data : s.Data.Concat(new byte[4096 - s.Data.Length]).ToArray())).ToArray();
        int total = streams.Sum(s => (s.Data.Length + 511) / 512);
        var file = new byte[512 * (3 + total)];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(file, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(24), 0x3E);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(26), 3);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(28), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(30), 9);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(32), 6);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(44), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(48), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(56), 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(60), 0xFFFFFFFE);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(68), 0xFFFFFFFE);
        for (int i = 0; i < 109; i++) BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(76 + 4 * i), i == 0 ? 0u : 0xFFFFFFFF);
        var fat = file.AsSpan(512, 512);
        for (int i = 0; i < 128; i++) BinaryPrimitives.WriteUInt32LittleEndian(fat[(4 * i)..], 0xFFFFFFFF);
        BinaryPrimitives.WriteUInt32LittleEndian(fat, 0xFFFFFFFD);
        BinaryPrimitives.WriteUInt32LittleEndian(fat[4..], 0xFFFFFFFE);

        void Entry(int index, string name, int type, uint right, uint child, uint start, long size)
        {
            var e = file.AsSpan(1024 + 128 * index, 128);
            var nameBytes = Encoding.Unicode.GetBytes(name);
            nameBytes.CopyTo(e);
            BinaryPrimitives.WriteUInt16LittleEndian(e[64..], (ushort)(nameBytes.Length + 2));
            e[66] = (byte)type;
            BinaryPrimitives.WriteUInt32LittleEndian(e[68..], 0xFFFFFFFF);
            BinaryPrimitives.WriteUInt32LittleEndian(e[72..], right);
            BinaryPrimitives.WriteUInt32LittleEndian(e[76..], child);
            BinaryPrimitives.WriteUInt32LittleEndian(e[116..], start);
            BinaryPrimitives.WriteInt64LittleEndian(e[120..], size);
        }
        Entry(0, "Root Entry", 5, 0xFFFFFFFF, 1, 0xFFFFFFFE, 0);
        int sector = 2;
        for (int s = 0; s < streams.Length; s++)
        {
            var data = streams[s].Data;
            int count = (data.Length + 511) / 512;
            for (int i = 0; i < count; i++) BinaryPrimitives.WriteUInt32LittleEndian(fat[(4 * (sector + i))..], i == count - 1 ? 0xFFFFFFFE : (uint)(sector + i + 1));
            Entry(1 + s, streams[s].Name, 2, s + 1 < streams.Length ? (uint)(s + 2) : 0xFFFFFFFF, 0xFFFFFFFF, (uint)sector, data.Length);
            data.CopyTo(file, 512 * (1 + sector));
            sector += count;
        }
        return file;
    }

    /// <summary>
    /// A minimal Word 97 document: a FIB with the given story lengths and (fc, lcb) pairs into the table stream, one compressed
    /// (8-bit) piece holding all the text at byte 1024 of the WordDocument stream, and the piece table at the start of the table stream.
    /// Character runs, if given, go in one CHPX FKP page after the text, each run's sprms covering its length of text in turn.
    /// </summary>
    private static byte[] Doc(string text, int[] storyLengths, params (int Pair, byte[] Data)[] tables) => Doc(text, storyLengths, [], tables);

    private static byte[] Doc(string text, int[] storyLengths, (int Length, byte[] Sprms)[] runs, params (int Pair, byte[] Data)[] tables)
    {
        int fkpPage = (1024 + text.Length + 511) / 512;
        var main = new byte[runs.Length > 0 ? 512 * (fkpPage + 1) : 1024 + text.Length];
        if (runs.Length > 0)
        {
            var fkp = main.AsSpan(512 * fkpPage, 512);
            int fc = 1024, at = 511;
            for (int r = 0; r < runs.Length; r++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(fkp[(4 * r)..], fc);
                fc += runs[r].Length;
                at = (at - runs[r].Sprms.Length - 1) & ~1;
                fkp[at] = (byte)runs[r].Sprms.Length;
                runs[r].Sprms.CopyTo(fkp[(at + 1)..]);
                fkp[4 * (runs.Length + 1) + r] = (byte)(at / 2);
            }
            BinaryPrimitives.WriteInt32LittleEndian(fkp[(4 * runs.Length)..], fc);
            fkp[511] = (byte)runs.Length;
            tables = [.. tables, (12, [.. BitConverter.GetBytes(1024), .. BitConverter.GetBytes(fc), .. BitConverter.GetBytes(fkpPage)])];
        }
        BinaryPrimitives.WriteUInt16LittleEndian(main, 0xA5EC);
        BinaryPrimitives.WriteUInt16LittleEndian(main.AsSpan(2), 0xC1);
        BinaryPrimitives.WriteUInt16LittleEndian(main.AsSpan(0x0A), 0x0200); // the table stream is 1Table
        BinaryPrimitives.WriteUInt16LittleEndian(main.AsSpan(32), 14);
        BinaryPrimitives.WriteUInt16LittleEndian(main.AsSpan(62), 22);
        for (int i = 0; i < storyLengths.Length; i++) BinaryPrimitives.WriteInt32LittleEndian(main.AsSpan(64 + 4 * (3 + i)), storyLengths[i]);
        BinaryPrimitives.WriteUInt16LittleEndian(main.AsSpan(152), 93);
        Encoding.Latin1.GetBytes(text).CopyTo(main, 1024);

        var table = new List<byte> { 2 };
        table.AddRange(BitConverter.GetBytes(4 + 4 + 8));
        table.AddRange(BitConverter.GetBytes(0));
        table.AddRange(BitConverter.GetBytes(text.Length));
        table.AddRange(BitConverter.GetBytes((ushort)0));
        table.AddRange(BitConverter.GetBytes((1024 * 2) | 0x40000000));
        table.AddRange(BitConverter.GetBytes((ushort)0));
        void Pair(int pair, int fc, int lcb)
        {
            BinaryPrimitives.WriteInt32LittleEndian(main.AsSpan(154 + 8 * pair), fc);
            BinaryPrimitives.WriteInt32LittleEndian(main.AsSpan(154 + 8 * pair + 4), lcb);
        }
        Pair(33, 0, table.Count);
        foreach (var (pair, data) in tables) { Pair(pair, table.Count, data.Length); table.AddRange(data); }
        return Cfb(("WordDocument", main), ("1Table", table.ToArray()));
    }

    private static string Text(IEnumerable<Block>? blocks) =>
        string.Join("|", (blocks ?? []).OfType<Paragraph>().Select(p => string.Concat(p.Inlines.Select(i => i switch { TextRun r => r.Text, FieldRun f => $"[{f.Kind}]", _ => "" }))));

    [Fact]
    public void Headers_and_footers_come_from_the_header_stories_with_live_page_fields()
    {
        const string body = "Body\r", header = "Head\r", footer = "Page \x13 PAGE \x14" + "1\x15\r";
        int f = header.Length + footer.Length;
        // Six separator stories, then even header, odd header, even footer, odd footer, first header and first footer, and a guard.
        int[] cps = [0, 0, 0, 0, 0, 0, 0, 0, header.Length, header.Length, f, f, f, f, f + 1];
        var plcfHdd = cps.SelectMany(BitConverter.GetBytes).ToArray();
        var doc = DocReader.Read(new MemoryStream(Doc(body + header + footer + "\r", [body.Length, 0, f + 1], (11, plcfHdd))));
        var section = Assert.Single(doc.Sections);
        Assert.Equal("Body", Text(section.Blocks));
        Assert.Equal("Head", Text(section.DefaultHeader?.Blocks));
        Assert.Equal("Page [Page]", Text(section.DefaultFooter?.Blocks));
        Assert.Null(section.FirstHeader);
    }

    [Fact]
    public void Caps_small_caps_and_hidden_text_use_their_own_sprms()
    {
        const string text = "Caps small gone\r";
        var doc = DocReader.Read(new MemoryStream(Doc(text, [text.Length], runs:
            [(5, [0x3B, 0x08, 1]), (6, [0x3A, 0x08, 1]), (4, [0x3C, 0x08, 1]), (1, [])])));
        var runs = doc.Sections[0].Blocks.OfType<Paragraph>().SelectMany(p => p.Inlines).OfType<TextRun>().ToList();
        Assert.Equal("Caps small ", string.Concat(runs.Select(r => r.Text)));
        Assert.True(runs[0].Properties.Caps && !runs[0].Properties.SmallCaps);
        Assert.True(runs[^1].Properties.SmallCaps && !runs[^1].Properties.Caps);
    }

    [Fact]
    public void A_group_floats_each_shape_at_its_place_in_the_anchor_frame()
    {
        static byte[] Rect(int l, int t, int r, int b) => [.. new[] { l, t, r, b }.SelectMany(BitConverter.GetBytes)];
        var group = Escher(0xF, 0xF004, Escher(0x1, 0xF009, Rect(0, 0, 1000, 1000)), Escher(2, 0xF00A, BitConverter.GetBytes(1025), BitConverter.GetBytes(0x201)));
        var child = Escher(0xF, 0xF004, Escher(1 << 4 | 2, 0xF00A, BitConverter.GetBytes(1026), BitConverter.GetBytes(0xA02)), Escher(0, 0xF00F, Rect(500, 0, 1000, 500)));
        var patriarch = Escher(0xF, 0xF004, Escher(2, 0xF00A, BitConverter.GetBytes(1024), BitConverter.GetBytes(0x5)));
        byte[] art = [.. Escher(0xF, 0xF000), 0, .. Escher(0xF, 0xF002, Escher(0xF, 0xF003, patriarch, Escher(0xF, 0xF003, group, child)))];
        byte[] spa = [.. new[] { 0, 2, 1025, 200, 400, 2200, 1400 }.SelectMany(BitConverter.GetBytes), 0, 0, 0, 0, 0, 0];
        var doc = DocReader.Read(new MemoryStream(Doc("\x08\r", [2], (40, spa), (50, art))));
        var box = Assert.Single(doc.Sections[0].Blocks.OfType<Paragraph>().SelectMany(p => p.Inlines).OfType<FloatingTextBox>());
        Assert.Equal((50, 25), (box.Width, box.Height));
        Assert.Equal((60, 20), (box.Horizontal.Offset, box.Vertical.Offset));
    }

    private static byte[] Escher(int verInst, int type, params byte[][] body)
    {
        var data = body.SelectMany(b => b).ToArray();
        return [.. BitConverter.GetBytes((ushort)verInst), .. BitConverter.GetBytes((ushort)type), .. BitConverter.GetBytes(data.Length), .. data];
    }

    [Fact]
    public void A_text_box_anchored_in_the_text_floats_with_its_story_and_position()
    {
        const string body = "\x08\r", inside = "Inside\r";
        var sp = Escher(0xF, 0xF004, Escher(202 << 4 | 2, 0xF00A, BitConverter.GetBytes(1026), BitConverter.GetBytes(0xA00)),
            Escher(2 << 4 | 3, 0xF00B, BitConverter.GetBytes((ushort)0x80), BitConverter.GetBytes(0x10000), BitConverter.GetBytes((ushort)0x390), BitConverter.GetBytes(1)));
        byte[] art = [.. Escher(0xF, 0xF000), 0, .. Escher(0xF, 0xF002, Escher(0xF, 0xF003, sp))];
        byte[] spa = [.. new[] { 0, 2, 1026, 200, 400, 2200, 1400 }.SelectMany(BitConverter.GetBytes), 0, 0, 0, 0, 0, 0];
        byte[] txbx = [.. new[] { 0, inside.Length, inside.Length + 1 }.SelectMany(BitConverter.GetBytes), .. new byte[44]];
        var doc = DocReader.Read(new MemoryStream(Doc(body + inside + "\r", [body.Length, 0, 0, 0, 0, 0, inside.Length + 1], (40, spa), (50, art), (56, txbx))));
        var box = Assert.Single(doc.Sections[0].Blocks.OfType<Paragraph>().SelectMany(p => p.Inlines).OfType<FloatingTextBox>());
        Assert.Equal((100, 50), (box.Width, box.Height));
        Assert.Equal(new FloatPosition(FloatReference.Page, FloatAlignment.Offset, 10), box.Horizontal);
        Assert.Equal("Inside", Text(box.Blocks));
    }
}
