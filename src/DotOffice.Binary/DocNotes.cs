using DotOffice.Core;
using DotOffice.Docx;

namespace DotOffice.Binary;

/// <summary>
/// Footnotes and endnotes of a .doc: PlcffndRef / PlcfendRef give the main-text positions of the references (each with an FRD whose
/// nAuto is 0 for a custom mark), PlcffndTxt / PlcfendTxt the start of each note's text in the footnote or endnote story. A note's text
/// starts with the 0x02 of its own mark.
/// </summary>
internal sealed partial class DocContext
{
    private const int MaxNotes = 100_000, MaxNoteDepth = 2;

    private readonly record struct NoteRef(NoteKind Kind, int Start, int End, bool Auto);

    private readonly Dictionary<int, NoteRef> _noteRefs = [];
    private Note? _currentNote;
    private int _noteDepth, _footnotes, _endnotes;

    private void ReadNotes()
    {
        int footnoteBase = _textLength;
        int endnoteBase = _textLength;
        for (int k = 4; k <= 7; k++) endnoteBase += Math.Max(Lw(k), 0); // footnotes, headers, macros (unused) and comments come first
        ReadNotes(NoteKind.Footnote, 2, 3, footnoteBase);
        ReadNotes(NoteKind.Endnote, 46, 47, endnoteBase);
    }

    private void ReadNotes(NoteKind kind, int refPair, int textPair, int storyBase)
    {
        int fc = FcLcb(refPair, out int lcb), tfc = FcLcb(textPair, out int tlcb);
        if (lcb < 10 || fc < 0 || (long)fc + lcb > table.Length || tlcb < 8 || tfc < 0 || (long)tfc + tlcb > table.Length) return;
        var refs = table.AsSpan(fc, lcb);
        var texts = table.AsSpan(tfc, tlcb);
        int n = Math.Min((lcb - 4) / 6, MaxNotes), textCount = tlcb / 4 - 1;
        for (int i = 0; i < n && i < textCount; i++)
        {
            long cp = U32(refs, 4 * i), start = storyBase + U32(texts, 4 * i), end = storyBase + U32(texts, 4 * (i + 1));
            bool auto = I16(refs, 4 * (n + 1) + 2 * i) != 0;
            if (cp < _textLength && start < end && end <= _text.Length) _noteRefs.TryAdd((int)cp, new NoteRef(kind, (int)start, (int)end, auto));
        }
    }

    /// <summary>The note referenced at a main-text position, built from its story and numbered; null when there is none.</summary>
    private Note? NoteAt(int cp, Document doc)
    {
        if (_noteDepth >= MaxNoteDepth || !_noteRefs.TryGetValue(cp, out var reference)) return null;
        // Footnotes are numbered 1, 2, 3 from the DOP's starting number (nFtn), endnotes i, ii, iii.
        var mark = !reference.Auto ? ""
            : reference.Kind == NoteKind.Footnote ? ListNumbers.Format(FootnoteStart() + _footnotes++, "decimal")
            : ListNumbers.Format(1 + _endnotes++, "lowerRoman");
        var note = new Note(reference.Kind) { Mark = mark };
        var outer = _currentNote;
        _currentNote = note;
        _noteDepth++;
        try
        {
            note.Blocks.AddRange(Build(reference.Start, reference.End, story: true).Sections[0].Blocks);
        }
        finally
        {
            _noteDepth--;
            _currentNote = outer;
        }
        if (reference.Kind == NoteKind.Endnote) doc.Endnotes.Add(note);
        return note;
    }

    private int FootnoteStart()
    {
        int fc = FcLcb(31, out int lcb);
        int start = lcb >= 4 && fc >= 0 && (long)fc + 4 <= table.Length ? U16(table, fc + 2) >> 2 : 1;
        return start > 0 ? start : 1;
    }
}
