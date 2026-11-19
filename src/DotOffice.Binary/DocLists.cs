using System.Text;
using DotOffice.Core;
using DotOffice.Docx;

namespace DotOffice.Binary;

/// <summary>
/// The lists of a .doc: list definitions (PlfLst, each an LSTF followed by its LVLs) and the list overrides that paragraphs name with
/// sprmPIlfo (PlfLfo, each an LFO with LFOLVLs that restart or reformat levels). A level's LVL holds the start number, the number format,
/// the label text with level placeholders (characters 0-8), what follows the label, and the label's paragraph and character sprms.
/// </summary>
internal sealed partial class DocContext
{
    private const int MaxLists = 10_000;

    private sealed class ListLevel
    {
        public int Start = 1, Nfc, Jc, Follow;
        public bool Legal, NoRestart;
        public string Text = "";
        public byte[] Papx = [], Chpx = [];
    }

    private sealed class ListOverride
    {
        public int Lsid;
        public readonly Dictionary<int, int> Starts = [];
        public readonly Dictionary<int, ListLevel> Levels = [];
    }

    private readonly Dictionary<int, ListLevel[]> _lists = [];
    private readonly List<ListOverride> _overrides = [];
    private readonly Dictionary<object, int?[]> _listCounters = [];

    private void ReadLists()
    {
        int fc = FcLcb(73, out int lcb);
        if (lcb >= 2 && fc >= 0 && fc < table.Length)
        {
            var d = table.AsSpan(fc);
            int count = Math.Min(I16(d, 0), MaxLists), at = 2;
            var definitions = new List<(int Lsid, bool Simple)>();
            for (int i = 0; i < count && at + 28 <= d.Length; i++, at += 28) definitions.Add(((int)U32(d, at), (d[at + 26] & 1) != 0));
            // The levels follow the LSTFs: one for a simple list, nine for any other.
            foreach (var (lsid, simple) in definitions)
            {
                var levels = new ListLevel[simple ? 1 : 9];
                for (int l = 0; l < levels.Length; l++)
                    if (ReadLevel(d, ref at) is { } level) levels[l] = level;
                    else return;
                _lists.TryAdd(lsid, levels);
            }
        }

        fc = FcLcb(74, out lcb);
        if (lcb < 4 || fc < 0 || (long)fc + lcb > table.Length) return;
        var o = table.AsSpan(fc, lcb);
        int lfos = (int)Math.Min(U32(o, 0), MaxLists), pos = 4;
        var counts = new List<int>();
        for (int i = 0; i < lfos && pos + 16 <= o.Length; i++, pos += 16)
        {
            _overrides.Add(new ListOverride { Lsid = (int)U32(o, pos) });
            counts.Add(o[pos + 12]);
        }
        // After the LFOs, each one's LFOData: a cp, then its LFOLVLs, each followed by an LVL when it replaces the level's formatting.
        for (int i = 0; i < counts.Count; i++)
        {
            if (pos + 4 > o.Length) return;
            pos += 4;
            for (int k = 0; k < counts[i]; k++)
            {
                if (pos + 8 > o.Length) return;
                int start = (int)U32(o, pos), flags = (int)U32(o, pos + 4), ilvl = flags & 0xF;
                pos += 8;
                if ((flags & 0x10) != 0 && ilvl < 9) _overrides[i].Starts[ilvl] = start;
                if ((flags & 0x20) != 0)
                {
                    if (ReadLevel(o, ref pos) is not { } level) return;
                    if (ilvl < 9) _overrides[i].Levels[ilvl] = level;
                }
            }
        }
    }

    /// <summary>An LVL: the 28-byte LVLF, its paragraph sprms, its character sprms and the label text (Xst).</summary>
    private static ListLevel? ReadLevel(ReadOnlySpan<byte> d, ref int at)
    {
        if (at + 28 > d.Length) return null;
        var level = new ListLevel
        {
            Start = (int)U32(d, at), Nfc = d[at + 4], Jc = d[at + 5] & 3, Legal = (d[at + 5] & 4) != 0, NoRestart = (d[at + 5] & 8) != 0, Follow = d[at + 15],
        };
        int cbChpx = d[at + 24], cbPapx = d[at + 25];
        at += 28;
        if (at + cbPapx + cbChpx + 2 > d.Length) return null;
        level.Papx = d.Slice(at, cbPapx).ToArray();
        at += cbPapx;
        level.Chpx = d.Slice(at, cbChpx).ToArray();
        at += cbChpx;
        int cch = U16(d, at);
        at += 2;
        if (at + 2 * cch > d.Length) return null;
        level.Text = Encoding.Unicode.GetString(d.Slice(at, 2 * cch));
        at += 2 * cch;
        return level;
    }

    /// <summary>The level a paragraph's list override (1-based) and level number point to, with the override's own formatting first.</summary>
    private ListLevel? LevelOf(int ilfo, int ilvl)
    {
        if (ilfo < 1 || ilfo > _overrides.Count || ilvl is < 0 or > 8) return null;
        var lfo = _overrides[ilfo - 1];
        if (lfo.Levels.TryGetValue(ilvl, out var own)) return own;
        return _lists.TryGetValue(lfo.Lsid, out var levels) ? levels[Math.Min(ilvl, levels.Length - 1)] : null;
    }

    private static string FormatName(int nfc) => nfc switch
    {
        1 => "upperRoman", 2 => "lowerRoman", 3 => "upperLetter", 4 => "lowerLetter", 5 => "ordinal", 22 => "decimalZero", 23 => "bullet", 255 => "none",
        _ => "decimal",
    };

    /// <summary>
    /// Counts a list paragraph and returns its label text: the level's text with each placeholder replaced by that level's current number.
    /// Lists that share a definition count together unless an override restarts a level; a level restarts after a higher one.
    /// </summary>
    private string? NextLabel(int ilfo, int ilvl)
    {
        if (LevelOf(ilfo, ilvl) is not { } level) return null;
        var lfo = _overrides[ilfo - 1];
        object key = lfo.Starts.Count > 0 ? ilfo : lfo.Lsid;
        if (!_listCounters.TryGetValue(key, out var counters)) _listCounters[key] = counters = new int?[9];
        int StartOf(int l) => lfo.Starts.TryGetValue(l, out var s) ? s : LevelOf(ilfo, l)?.Start ?? 1;

        counters[ilvl] = counters[ilvl] is { } c ? c + 1 : StartOf(ilvl);
        for (int deeper = ilvl + 1; deeper < 9; deeper++)
            if (LevelOf(ilfo, deeper) is not { NoRestart: true }) counters[deeper] = null;

        if (level.Nfc == 23) return level.Text;
        var text = new StringBuilder();
        foreach (char ch in level.Text)
        {
            if (ch >= 9) { text.Append(ch); continue; }
            int referenced = ch;
            int value = counters[referenced] ?? StartOf(referenced);
            var format = level.Legal && referenced < ilvl ? "decimal" : FormatName(LevelOf(ilfo, referenced)?.Nfc ?? 0);
            text.Append(ListNumbers.Format(value, format));
        }
        return text.ToString();
    }

    private static ListLabelSuffix SuffixOf(ListLevel level) => level.Follow switch { 1 => ListLabelSuffix.Space, 2 => ListLabelSuffix.Nothing, _ => ListLabelSuffix.Tab };

    private static ParagraphAlignment LabelAlignment(ListLevel level) =>
        level.Jc switch { 1 => ParagraphAlignment.Center, 2 => ParagraphAlignment.Right, _ => ParagraphAlignment.Left };
}
