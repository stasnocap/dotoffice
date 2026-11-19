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
    /// Character runs, if given, go in one CHPX FKP page after the text, each run's sprms covering its length of text in turn; a Data
    /// stream, if given, goes in the file too.
    /// </summary>
    private static byte[] Doc(string text, int[] storyLengths, params (int Pair, byte[] Data)[] tables) => Doc(text, storyLengths, [], null, tables);

    private static byte[] Doc(string text, int[] storyLengths, (int Length, byte[] Sprms)[] runs, byte[]? data = null, params (int Pair, byte[] Data)[] tables) =>
        Doc(text, storyLengths, runs, [], data, tables);

    private static byte[] Doc(string text, int[] storyLengths, (int Length, byte[] Sprms)[] runs, (int Length, byte[] Sprms)[] paras, byte[]? data = null,
        params (int Pair, byte[] Data)[] tables)
    {
        int fkpPage = (1024 + text.Length + 511) / 512;
        int pages = (runs.Length > 0 ? 1 : 0) + (paras.Length > 0 ? 1 : 0);
        var main = new byte[pages > 0 ? 512 * (fkpPage + pages) : 1024 + text.Length];
        // One FKP page: the runs' fcs, then per run its offset (CHPX) or a BX (offset and 12-byte PHE, PAPX), and the property sets from the end.
        void Fkp(int page, (int Length, byte[] Sprms)[] items, bool papx, int pair)
        {
            var fkp = main.AsSpan(512 * page, 512);
            int fc = 1024, at = 511, entry = papx ? 13 : 1;
            for (int r = 0; r < items.Length; r++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(fkp[(4 * r)..], fc);
                fc += items[r].Length;
                // a PAPX is cb (or 0 and cb') and the istd (0) with the sprms, 2 * cb - 1 (or 2 * cb') bytes long
                byte[] props = papx ? [.. (byte[])(items[r].Sprms.Length % 2 == 1 ? [(byte)((items[r].Sprms.Length + 3) / 2)] : [0, (byte)((items[r].Sprms.Length + 2) / 2)]), 0, 0, .. items[r].Sprms]
                    : [(byte)items[r].Sprms.Length, .. items[r].Sprms];
                at = (at - props.Length) & ~1;
                props.CopyTo(fkp[at..]);
                fkp[4 * (items.Length + 1) + entry * r] = (byte)(at / 2);
            }
            BinaryPrimitives.WriteInt32LittleEndian(fkp[(4 * items.Length)..], fc);
            fkp[511] = (byte)items.Length;
            tables = [.. tables, (pair, [.. BitConverter.GetBytes(1024), .. BitConverter.GetBytes(fc), .. BitConverter.GetBytes(page)])];
        }
        if (runs.Length > 0) Fkp(fkpPage, runs, false, 12);
        if (paras.Length > 0) Fkp(fkpPage + pages - 1, paras, true, 13);
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
        foreach (var (pair, bytes) in tables) { Pair(pair, table.Count, bytes.Length); table.AddRange(bytes); }
        return data is null ? Cfb(("WordDocument", main), ("1Table", table.ToArray())) : Cfb(("WordDocument", main), ("1Table", table.ToArray()), ("Data", data));
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

    [Fact]
    public void A_picture_character_shows_the_picture_its_PICF_points_to_in_the_Data_stream()
    {
        var png = PngTestEncoder.Encode(2, 2, 2, 8, new int[12]);
        var shape = Escher(0xF, 0xF004, Escher(75 << 4 | 2, 0xF00A, BitConverter.GetBytes(1025), BitConverter.GetBytes(0xA00)),
            Escher(1 << 4 | 3, 0xF00B, BitConverter.GetBytes((ushort)0x4104), BitConverter.GetBytes(3))); // pib numbers the document's store
        var bse = Escher(6 << 4 | 2, 0xF007, new byte[36], Escher(0x6E0, 0xF01E, new byte[17], png));
        var picf = new byte[0x44];
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(4), 0x44);
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(6), 0x66); // a picture name follows the header
        foreach (var (at, value) in new[] { (28, 2000), (30, 1000), (32, 1000), (34, 500) }) BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(at), (ushort)value);
        byte[] picture = [.. new byte[16], .. picf, 1, (byte)'x', .. shape, .. bse];
        BinaryPrimitives.WriteInt32LittleEndian(picture.AsSpan(16), picture.Length - 16);
        var doc = DocReader.Read(new MemoryStream(Doc("\x01\r", [2], runs: [(1, [0x03, 0x6A, 16, 0, 0, 0, 0x55, 0x08, 1]), (1, [])], data: picture)));
        var image = Assert.Single(doc.Sections[0].Blocks.OfType<Paragraph>().SelectMany(p => p.Inlines).OfType<ImageRun>());
        Assert.Equal((100, 25), (image.Width, image.Height));
        Assert.Equal(png, image.Image.Data);
    }

    /// <summary>A font table (SttbfFfn) of FFNs: the family in the second byte, the alternative name's index at byte 5, the names at 40.</summary>
    private static byte[] Fonts(params (string Name, string? Alt, int Family)[] fonts)
    {
        var table = new List<byte>();
        table.AddRange(BitConverter.GetBytes((ushort)fonts.Length));
        table.AddRange(BitConverter.GetBytes((ushort)0));
        foreach (var (name, alt, family) in fonts)
        {
            var names = Encoding.Unicode.GetBytes(name + "\0" + (alt is null ? "" : alt + "\0"));
            var ffn = new byte[40 + names.Length];
            ffn[0] = (byte)(ffn.Length - 1);
            ffn[1] = (byte)(family << 4 | 2);
            ffn[5] = (byte)(alt is null ? 0 : name.Length + 1);
            names.CopyTo(ffn, 40);
            table.AddRange(ffn);
        }
        return [.. table];
    }

    [Fact]
    public void The_font_table_declares_each_fonts_family_and_alternative_name()
    {
        var doc = DocReader.Read(new MemoryStream(Doc("x\r", [2], (15, Fonts(("Times New Roman", null, 1), ("PragmaticaWINCTT", null, 2), ("Monotype Sorts", "Courier10 BT", 0))))));
        Assert.Equal(new FontDeclaration(null, true), doc.Fonts!["Times New Roman"]);
        Assert.Equal(new FontDeclaration(null, false), doc.Fonts["pragmaticawinctt"]);
        Assert.Equal(new FontDeclaration("Courier10 BT", false), doc.Fonts["Monotype Sorts"]);
    }

    [Fact]
    public void Paragraph_borders_come_from_both_border_sprms_and_doubles_are_three_widths()
    {
        const string text = "Old\rNew\r";
        byte[] old80 = [0x26, 0x64, 12, 3, 0, 13]; // sprmPBrcBottom80: width 12/8 pt, double, auto, 13 pt space
        byte[] brc = [0x4E, 0xC6, 8, 0xFF, 0, 0, 0, 4, 1, 4, 0]; // sprmPBrcTop: red, 1/2 pt, single, 4 pt space
        var doc = DocReader.Read(new MemoryStream(Doc(text, [text.Length], [], paras: [(4, old80), (4, brc)])));
        var paragraphs = doc.Sections[0].Blocks.OfType<Paragraph>().ToList();
        Assert.Equal(new Border(BorderStyle.Double, 4.5, Color.Black, 13), paragraphs[0].Properties.Borders.Bottom);
        Assert.Equal(new Border(BorderStyle.Single, 0.5, new Color(255, 0, 0), 4), paragraphs[1].Properties.Borders.Top);
        Assert.False(paragraphs[1].Properties.Borders.Bottom.IsVisible);
    }

    [Fact]
    public void Tab_stops_come_from_both_tab_sprms_measured_from_the_margin()
    {
        const string text = "A\tB\r";
        byte[] left = [0x0F, 0x84, 0xAA, 0x01]; // a 426-twip left indent does not move the tabs
        byte[] papx = [0x0D, 0xC6, 5, 0, 1, 0x96, 0x25, 0x0A]; // sprmPChgTabsPapx: add a right tab at 9622 twips with a dot leader
        byte[] chg = [0x15, 0xC6, 9, 1, 0x96, 0x25, 0x28, 0, 1, 0xA0, 0x05, 0]; // sprmPChgTabs: delete it (40-twip tolerance), add a left tab at 1440
        var doc = DocReader.Read(new MemoryStream(Doc(text, [text.Length], [], paras: [(text.Length, [.. left, .. papx])])));
        var tab = Assert.Single(doc.Sections[0].Blocks.OfType<Paragraph>().Single().Properties.TabStops);
        Assert.Equal(new TabStop(481.1, TabAlignment.Right, TabLeader.Dot), tab);
        doc = DocReader.Read(new MemoryStream(Doc(text, [text.Length], [], paras: [(text.Length, [.. papx, .. chg])])));
        Assert.Equal(new TabStop(72, TabAlignment.Left), Assert.Single(doc.Sections[0].Blocks.OfType<Paragraph>().Single().Properties.TabStops));
    }

    [Fact]
    public void Frame_sprms_float_a_paragraph_and_a_framed_header_paragraph_does_not_push_the_body_down()
    {
        const string body = "Body\r", header = "\x13 PAGE \x14" + "1\x15\r\r\r";
        // centered on the margin, one twip below the paragraph, wrapped around
        byte[] frame = [0x18, 0x84, 0xFC, 0xFF, 0x19, 0x84, 1, 0, 0x1B, 0x26, 0x60, 0x23, 0x24, 2];
        int[] cps = [0, 0, 0, 0, 0, 0, 0, 0, header.Length, header.Length, header.Length, header.Length, header.Length, header.Length + 1];
        var file = Doc(body + header + "\r", [body.Length, 0, header.Length + 1], [], paras: [(body.Length, []), (header.Length - 2, frame), (3, [])],
            tables: (11, cps.SelectMany(BitConverter.GetBytes).ToArray()));
        var doc = DocReader.Read(new MemoryStream(file));
        var blocks = doc.Sections[0].DefaultHeader!.Blocks.OfType<Paragraph>().ToList();
        Assert.Equal(2, blocks.Count); // the story's guard paragraph mark is left out
        var f = blocks[0].Properties.Frame!;
        Assert.Equal(new FloatPosition(FloatReference.Margin, FloatAlignment.Center, 0), f.Horizontal);
        Assert.Equal(new FloatPosition(FloatReference.Paragraph, FloatAlignment.Offset, 0.05), f.Vertical);
        Assert.Null(blocks[1].Properties.Frame);
        Assert.False(doc.Settings.PageBreakBeforeDropsSpaceBefore);

        var page = Layout.LayoutTestHelpers.Render(doc).Pages[0];
        var bodyText = Layout.LayoutTestHelpers.Texts(page).First(t => t.Text.StartsWith("Body"));
        Assert.True(bodyText.Baseline < doc.Sections[0].Margins.Top + 20, $"body baseline {bodyText.Baseline}");
    }

    /// <summary>An LVL: the LVLF (start, number format, what follows, the sprm lengths), the paragraph and character sprms, the label text.</summary>
    private static byte[] Lvl(int start, int nfc, string text, byte[] papx, byte[]? chpx = null, int follow = 0)
    {
        var lvlf = new byte[28];
        BinaryPrimitives.WriteInt32LittleEndian(lvlf, start);
        lvlf[4] = (byte)nfc;
        lvlf[15] = (byte)follow;
        lvlf[24] = (byte)(chpx?.Length ?? 0);
        lvlf[25] = (byte)papx.Length;
        return [.. lvlf, .. papx, .. chpx ?? [], .. BitConverter.GetBytes((ushort)text.Length), .. Encoding.Unicode.GetBytes(text)];
    }

    [Fact]
    public void List_paragraphs_get_numbers_and_bullets_from_their_list_levels()
    {
        const string text = "One\rTwo\rSub\rThree\rDot\r";
        byte[] Lstf(int lsid, bool simple) => [.. BitConverter.GetBytes(lsid), .. new byte[22], (byte)(simple ? 1 : 0), 0];
        byte[] hanging = [0x0F, 0x84, 0xD0, 0x02, 0x11, 0x84, 0x70, 0xFE]; // left 720, first line -400
        var numbered = Enumerable.Range(0, 9).SelectMany(l => Lvl(1, l == 1 ? 4 : 0, l == 1 ? "\0.\u0001)" : "\0.", hanging)).ToArray();
        byte[] plfLst = [.. BitConverter.GetBytes((short)2), .. Lstf(7, false), .. Lstf(9, true), .. numbered,
            .. Lvl(1, 23, "\uF0B7", [], [0x4F, 0x4A, 1, 0], follow: 1)]; // a Symbol bullet followed by a space
        byte[] Lfo(int lsid) => [.. BitConverter.GetBytes(lsid), .. new byte[12]];
        byte[] plfLfo = [.. BitConverter.GetBytes(2), .. Lfo(7), .. Lfo(9), .. new byte[8]];
        byte[] Item(int ilfo, int ilvl) => [0x0B, 0x46, (byte)ilfo, 0, 0x0A, 0x26, (byte)ilvl];
        var fonts = Fonts(("Times New Roman", null, 1), ("Symbol", null, 1));
        var doc = DocReader.Read(new MemoryStream(Doc(text, [text.Length], [], paras: [(4, Item(1, 0)), (4, Item(1, 0)), (4, Item(1, 1)), (6, Item(1, 0)), (4, Item(2, 0))],
            tables: [(15, fonts), (73, plfLst), (74, plfLfo)])));
        var paragraphs = doc.Sections[0].Blocks.OfType<Paragraph>().ToList();
        Assert.Equal(["1.", "2.", "2.a)", "3.", "•"], paragraphs.Select(p => p.ListLabel?.Text));
        Assert.Equal((36.0, -20.0), (paragraphs[0].Properties.LeftIndent, paragraphs[0].Properties.FirstLineIndent));
        Assert.Equal(ListLabelSuffix.Space, paragraphs[4].ListLabel!.Suffix);
        Assert.Equal("Times New Roman", paragraphs[4].ListLabel!.Properties.FontFamily);
    }

    [Fact]
    public void A_character_style_adds_its_formatting_to_the_paragraph_styles()
    {
        // style 0: Normal, bold, 12 pt; style 10: Default Paragraph Font; style 22: Hyperlink, underlined
        static byte[] Std(int sgc, int istdBase, string name, byte[][] upx)
        {
            var body = new List<byte>();
            body.AddRange(BitConverter.GetBytes((ushort)0));
            body.AddRange(BitConverter.GetBytes((ushort)(sgc | istdBase << 4)));
            body.AddRange(BitConverter.GetBytes((ushort)upx.Length));
            body.AddRange(new byte[4]);
            body.AddRange(BitConverter.GetBytes((ushort)name.Length));
            body.AddRange(Encoding.Unicode.GetBytes(name + "\0"));
            foreach (var u in upx)
            {
                if ((body.Count & 1) != 0) body.Add(0);
                body.AddRange(BitConverter.GetBytes((ushort)u.Length));
                body.AddRange(u);
            }
            return [.. BitConverter.GetBytes((ushort)body.Count), .. body];
        }
        var stsh = new List<byte>();
        stsh.AddRange(BitConverter.GetBytes((ushort)18));
        stsh.AddRange(BitConverter.GetBytes((ushort)23));
        stsh.AddRange(BitConverter.GetBytes((ushort)10));
        stsh.AddRange(new byte[14]);
        for (int i = 0; i < 23; i++)
            stsh.AddRange(i switch
            {
                0 => Std(1, 0xFFF, "Normal", [[0, 0], [0x35, 0x08, 1, 0x43, 0x4A, 24, 0]]),
                10 => Std(2, 0xFFF, "Default Paragraph Font", [[]]),
                22 => Std(2, 10, "Hyperlink", [[0x3E, 0x2A, 1]]),
                _ => BitConverter.GetBytes((ushort)0),
            });
        const string text = "Link\r";
        var doc = DocReader.Read(new MemoryStream(Doc(text, [text.Length], runs: [(4, [0x30, 0x4A, 22, 0]), (1, [])], tables: (1, stsh.ToArray()))));
        var run = doc.Sections[0].Blocks.OfType<Paragraph>().Single().Inlines.OfType<TextRun>().Single();
        Assert.True(run.Properties.Bold);
        Assert.Equal(12, run.Properties.FontSize);
        Assert.Equal(UnderlineStyle.Single, run.Properties.Underline);
    }

    [Fact]
    public void Footnotes_are_numbered_and_their_text_starts_with_their_mark()
    {
        const string body = "Text\x02 more\x02\r", notes = "\x02 First\r\x02 Second\r";
        byte[] fndRef = [.. new[] { 4, 10, 12 }.SelectMany(BitConverter.GetBytes), .. BitConverter.GetBytes((short)1), .. BitConverter.GetBytes((short)1)];
        byte[] fndTxt = [.. new[] { 0, 8, notes.Length, notes.Length + 1 }.SelectMany(BitConverter.GetBytes)];
        var doc = DocReader.Read(new MemoryStream(Doc(body + notes + "\r", [body.Length, notes.Length + 1], (2, fndRef), (3, fndTxt))));
        var references = doc.Sections[0].Blocks.OfType<Paragraph>().Single().Inlines.OfType<NoteReference>().ToList();
        Assert.Equal(["1", "2"], references.Select(r => r.Mark));
        var first = references[0].Note.Blocks.OfType<Paragraph>().Single();
        Assert.IsType<NoteMark>(first.Inlines[0]);
        Assert.Equal(" First", Assert.IsType<TextRun>(first.Inlines[1]).Text);
        Assert.NotNull(doc.NoteSeparator); // no separator story: Word's own, an empty Normal paragraph
    }

    [Fact]
    public void Html_paragraph_auto_spacing_collapses_the_space_between_paragraphs()
    {
        const string text = "One\rTwo\r";
        var dop = new byte[544];
        byte[] spaced = [0x13, 0xA4, 120, 0, 0x14, 0xA4, 120, 0]; // 6 pt before and after
        byte[] File(byte copts) { dop[512] = copts; return Doc(text, [text.Length], [], paras: [(4, spaced), (4, spaced)], tables: (31, dop)); }
        Assert.True(DocReader.Read(new MemoryStream(File(0))).Settings.CollapseParagraphSpacing);
        Assert.False(DocReader.Read(new MemoryStream(File(4))).Settings.CollapseParagraphSpacing); // fDontUseHTMLParagraphAutoSpacing

        double Gap(byte copts)
        {
            var page = Layout.LayoutTestHelpers.Render(DocReader.Read(new MemoryStream(File(copts)))).Pages[0];
            var lines = Layout.LayoutTestHelpers.Texts(page).ToList();
            return lines.First(t => t.Text.StartsWith("Two")).Baseline - lines.First(t => t.Text.StartsWith("One")).Baseline;
        }
        Assert.Equal(6, Gap(4) - Gap(0), 3);
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
