using System.Globalization;
using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>
/// Lays out a <see cref="Document"/> into pages: paragraphs are broken into lines, blocks become
/// slices, and slices are paginated honoring page breaks, keep-with-next, keep-lines-together,
/// widow/orphan control, row splitting and repeated header rows. Headers and footers are laid out
/// per page once the total page count is known.
/// </summary>
public sealed class LayoutEngine
{
    private const double Epsilon = 0.01;

    /// <summary>Space taken by the line separating footnotes from the body text.</summary>
    private const double NoteSeparatorHeight = 12;

    private readonly LayoutOptions _options;
    private readonly FontSelector _fonts;
    private readonly Itemizer _itemizer;

    public LayoutEngine(LayoutOptions options)
    {
        _options = options;
        _fonts = new FontSelector(options.Fonts, options.Warnings);
        _itemizer = new Itemizer(_fonts);
    }

    /// <summary>A slice on a page: <see cref="Dx"/> moves it from the column it was built for to its column.</summary>
    private readonly record struct Placement(Slice Slice, double Dx, double Y, Region Region, int Column);

    /// <summary>A horizontal band of a page laid out in the columns of one section.</summary>
    private sealed class Region(Section section, double top, IReadOnlyList<(double X, double Width)> columns)
    {
        public Section Section { get; } = section;
        public double Top { get; } = top;

        /// <summary>Absolute x and width of each column.</summary>
        public IReadOnlyList<(double X, double Width)> Columns { get; } = columns;

        /// <summary>Lowest point reached by content in any column.</summary>
        public double Bottom { get; set; } = top;
    }

    private sealed class Page
    {
        public required Section Section { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
        public double ContentTop { get; set; }
        public double ContentBottom { get; set; }
        public HeaderFooter? Header { get; init; }
        public HeaderFooter? Footer { get; init; }
        public int Number { get; init; }
        public List<Placement> Placed { get; } = [];
        public List<Region> Regions { get; } = [];
        public Region Region => Regions[^1];

        /// <summary>Current column of <see cref="Region"/> and the index of its first placement.</summary>
        public int Column { get; set; }
        public int ColumnStart { get; set; }

        /// <summary>The current column was started because content overflowed (not by a forced break or section start).</summary>
        public bool NaturalBreak { get; set; }
        public double Cursor { get; set; }
        public bool HasBody => Placed.Exists(p => !IsRepeated(p.Slice));
        public bool ColumnHasBody => Placed.Skip(ColumnStart).Any(p => !IsRepeated(p.Slice));
        public int ColumnCount => Placed.Count - ColumnStart;
        public HashSet<Slice> Repeated { get; } = [];
        public bool IsRepeated(Slice s) => Repeated.Contains(s);

        /// <summary>Footnote lines at the page bottom, with the body slice that referenced them (null when carried over).</summary>
        public List<(Slice Slice, Slice? Owner)> Notes { get; } = [];

        /// <summary>The first footnote lines continue a footnote from the previous page.</summary>
        public bool NotesContinued { get; set; }

        public double NotesHeight => Notes.Count == 0 ? 0 : NoteSeparatorHeight + Notes.Sum(n => n.Slice.Height + n.Slice.TrailingSpace);

        /// <summary>Bottom of the body text: the content area minus footnotes.</summary>
        public double BodyBottom => ContentBottom - NotesHeight;

        /// <summary>Floating images, painted below (<c>Behind</c>) or above the text.</summary>
        public List<(ImageItem Item, bool Behind)> Floats { get; } = [];

        /// <summary>Areas wrapped text avoids.</summary>
        public List<Exclusion> Exclusions { get; } = [];

        public void StartRegion(Section section, double top)
        {
            var columns = section.Columns.Resolve(section.ContentWidth).Select(c => (section.Margins.Left + c.X, c.Width)).ToList();
            Regions.Add(new Region(section, top, columns));
            Column = 0;
            ColumnStart = Placed.Count;
            Cursor = top;
            NaturalBreak = false;
        }
    }

    /// <summary>Page number and number format of a page, as seen by PAGEREF fields.</summary>
    private readonly record struct PageLabel(int Number, PageNumberFormat Format);

    public LayoutDocument Layout(Document document, CancellationToken cancellationToken = default)
    {
        // PAGEREF results are only known after pagination; when the document has any, it is laid out
        // a second time with the page numbers from the first pass so that their widths are right.
        var pages = Paginate(document, null, out bool usesPageRefs, cancellationToken);
        bool renumbered = NumberFootnotesPerPage(pages);
        if (usesPageRefs || renumbered) pages = Paginate(document, Bookmarks(pages), out _, cancellationToken);

        var settings = document.Settings;
        var fields = new FieldValues(pages.Count, Bookmarks(pages),
            pages.GroupBy(p => p.Section).ToDictionary(g => g.Key, g => g.Count()));
        var result = new LayoutDocument { Title = document.Title, Author = document.Author };
        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            result.Pages.Add(Render(page, fields, settings));
        }
        return result;
    }

    private List<Page> Paginate(Document document, Dictionary<string, PageLabel>? bookmarks, out bool usesPageRefs,
        CancellationToken cancellationToken)
    {
        var pages = new List<Page>();
        var settings = document.Settings;
        bool pageRefs = false;
        var bodyContext = Context(settings, field =>
        {
            // Page numbers are only known after pagination: reserve two digits, which fits most documents.
            if (field.Kind != FieldKind.PageRef) return "99";
            pageRefs = true;
            return bookmarks is not null && field.Bookmark is not null && bookmarks.TryGetValue(field.Bookmark, out var label)
                ? Format(label.Number, field.NumberFormat ?? label.Format)
                : field.CachedResult is { Length: > 0 } cached ? cached : "1";
        });

        Page? current = null;
        Paginator? last = null;
        var carry = new List<Slice>(); // footnote lines continued on the next page
        int pageNumber = 0;
        Section? numberingSection = null;
        foreach (var section in document.Sections)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool sameSize = current is not null && Math.Abs(current.Width - section.PageWidth) < 0.5
                                                && Math.Abs(current.Height - section.PageHeight) < 0.5;
            if (current is null || section.Start != SectionStart.Continuous || !sameSize)
            {
                if (current is not null && section.Start is SectionStart.EvenPage or SectionStart.OddPage)
                {
                    bool nextIsEven = (pages.Count + 1) % 2 == 0;
                    if (nextIsEven != (section.Start == SectionStart.EvenPage))
                        pages.Add(current = NewPage(section, NextNumber(section, ref numberingSection, ref pageNumber), false, settings));
                }
                current = NewPage(section, NextNumber(section, ref numberingSection, ref pageNumber), true, settings);
                pages.Add(current);
            }

            var notes = new Dictionary<Note, List<Slice>>(ReferenceEqualityComparer.Instance);
            var paginator = new Paginator(pages, () =>
            {
                var page = NewPage(section, NextNumber(section, ref numberingSection, ref pageNumber), false, settings);
                current = page;
                return page;
            }, note =>
            {
                if (!notes.TryGetValue(note, out var lines))
                    notes[note] = lines = BlockFlow.Build(note.Blocks, section.Margins.Left, section.ContentWidth, bodyContext);
                return lines;
            }, carry);
            last = paginator;
            paginator.StartSection(section);
            bool keepWithNext = false;
            for (int i = 0; i < section.Blocks.Count; i++)
            {
                // Blocks are built for the column they start in, so columns of different widths work.
                var (x, width) = paginator.Column;
                bool keep = keepWithNext;
                var floats = section.Blocks[i] is Paragraph p ? p.Inlines.OfType<FloatingImage>().Where(f => !f.Properties.Hidden).ToList() : [];
                var slices = BlockFlow.BuildBlock(section.Blocks, i, x, width, bodyContext, ref keepWithNext, paginator.Obstacles);
                if ((floats.Count > 0 || section.Columns.Definitions.Count > 1) && slices.Count > 0 && paginator.StartsInNextColumn(slices[0]))
                    paginator.NextColumn();
                if (floats.Count > 0)
                {
                    // The anchor paragraph wraps around its own floats, so they are positioned before it is laid out.
                    paginator.PlaceFloats(floats);
                }
                if (floats.Count > 0 || paginator.Column.X != x)
                {
                    (x, width) = paginator.Column;
                    keepWithNext = keep;
                    slices = BlockFlow.BuildBlock(section.Blocks, i, x, width, bodyContext, ref keepWithNext, paginator.Obstacles);
                }
                foreach (var slice in slices) paginator.Add(slice, x);
            }
            int index = document.Sections.IndexOf(section);
            var next = index + 1 < document.Sections.Count ? document.Sections[index + 1] : null;
            // A continuous section break balances the columns above it, as in Word.
            if (next is { Start: SectionStart.Continuous } && Math.Abs(next.PageWidth - section.PageWidth) < 0.5
                                                        && Math.Abs(next.PageHeight - section.PageHeight) < 0.5)
                paginator.Balance();
            current = pages[^1];
        }
        if (last is not null && document.Endnotes.Count > 0)
        {
            // Endnotes follow the text of the last section, after a short separator line.
            var (x, width) = last.Column;
            var separator = new Slice { Height = NoteSeparatorHeight, LeadingSpace = NoteSeparatorHeight };
            separator.Items.Add(new LineItem(x, NoteSeparatorHeight / 2, x + Math.Min(144, width / 3), NoteSeparatorHeight / 2, 0.5, Color.Black));
            last.Add(separator, x);
            var blocks = document.Endnotes.SelectMany(n => n.Blocks).ToList();
            bool keep = false;
            for (int i = 0; i < blocks.Count; i++)
                foreach (var slice in BlockFlow.BuildBlock(blocks, i, x, width, bodyContext, ref keep))
                    last.Add(slice, x);
        }
        last?.FlushNotes();
        usesPageRefs = pageRefs;
        return pages;
    }

    /// <summary>
    /// Gives footnotes numbered per page (<see cref="Note.PageMark"/>) their number on the page they were
    /// placed on; returns true when a mark changed, so the document must be laid out again.
    /// </summary>
    private static bool NumberFootnotesPerPage(List<Page> pages)
    {
        bool changed = false;
        foreach (var page in pages)
        {
            int n = 0;
            var seen = new HashSet<Note>(ReferenceEqualityComparer.Instance);
            foreach (var note in page.Placed.SelectMany(p => p.Slice.Footnotes))
            {
                if (note.PageMark is null || !seen.Add(note)) continue;
                var mark = note.PageMark(++n);
                if (mark != note.Mark) (note.Mark, changed) = (mark, true);
            }
        }
        return changed;
    }

    /// <summary>The page of every bookmark in the body (first occurrence wins).</summary>
    private static Dictionary<string, PageLabel> Bookmarks(List<Page> pages)
    {
        var result = new Dictionary<string, PageLabel>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in pages)
            foreach (var placement in page.Placed)
                foreach (var item in placement.Slice.Items)
                    if (item is AnchorItem anchor)
                        result.TryAdd(anchor.Name, new PageLabel(page.Number, page.Section.PageNumberFormat));
        return result;
    }

    /// <summary>Values of page-dependent fields once pagination is final.</summary>
    private sealed class FieldValues(int total, Dictionary<string, PageLabel> bookmarks, Dictionary<Section, int> sectionPages)
    {
        public string Text(FieldRun field, Page page) => field.Kind switch
        {
            FieldKind.Page => Format(page.Number, field.NumberFormat ?? page.Section.PageNumberFormat),
            FieldKind.NumPages => Format(total, field.NumberFormat ?? PageNumberFormat.Decimal),
            FieldKind.SectionPages => Format(sectionPages.GetValueOrDefault(page.Section, 1), field.NumberFormat ?? PageNumberFormat.Decimal),
            FieldKind.PageRef when field.Bookmark is not null && bookmarks.TryGetValue(field.Bookmark, out var label) =>
                Format(label.Number, field.NumberFormat ?? label.Format),
            _ => field.CachedResult is { Length: > 0 } cached ? cached : "Error! Bookmark not defined.",
        };
    }

    private static int NextNumber(Section section, ref Section? numberingSection, ref int number)
    {
        if (!ReferenceEquals(section, numberingSection))
        {
            numberingSection = section;
            if (section.PageNumberStart is { } start) return number = start;
        }
        return ++number;
    }

    private FlowContext Context(DocumentSettings settings, Func<FieldRun, string> fieldText) => new()
    {
        Fonts = _fonts, Itemizer = _itemizer, Settings = settings, FieldText = fieldText, Warnings = _options.Warnings,
    };

    private Page NewPage(Section section, int number, bool sectionStart, DocumentSettings settings)
    {
        bool first = sectionStart && section.DifferentFirstPage;
        bool even = settings.EvenAndOddHeaders && number % 2 == 0;
        var header = first ? section.FirstHeader : even ? section.EvenHeader : section.DefaultHeader;
        var footer = first ? section.FirstFooter : even ? section.EvenFooter : section.DefaultFooter;

        double top = section.Margins.Top, bottom = section.PageHeight - section.Margins.Bottom;
        var placeholder = Context(settings, field => field.Kind == FieldKind.Page
            ? Format(number, field.NumberFormat ?? section.PageNumberFormat)
            : field.Kind == FieldKind.PageRef ? field.CachedResult ?? "99" : "99");
        if (header is not null)
        {
            double h = BlockFlow.Stack(BlockFlow.Build(header.Blocks, section.Margins.Left, section.ContentWidth, placeholder), []);
            top = Math.Max(top, section.HeaderDistance + h);
        }
        if (footer is not null)
        {
            double h = BlockFlow.Stack(BlockFlow.Build(footer.Blocks, section.Margins.Left, section.ContentWidth, placeholder), []);
            bottom = Math.Min(bottom, section.PageHeight - section.FooterDistance - h);
        }
        if (bottom - top < 36) bottom = top + 36; // absurd margins: keep a usable body area

        var page = new Page
        {
            Section = section, Width = section.PageWidth, Height = section.PageHeight,
            ContentTop = top, ContentBottom = bottom,
            Header = header, Footer = footer, Number = number,
        };
        page.StartRegion(section, top);
        return page;
    }

    private LayoutPage Render(Page page, FieldValues fields, DocumentSettings settings)
    {
        var output = new LayoutPage(page.Width, page.Height);
        var section = page.Section;
        var context = Context(settings, field => fields.Text(field, page));
        var front = page.Floats.Where(f => !f.Behind).Select(f => f.Item).ToList();
        output.Items.AddRange(page.Floats.Where(f => f.Behind).Select(f => f.Item));

        void Story(List<(double Y, Slice Slice)> stacked, double top)
        {
            foreach (var (y, slice) in stacked)
            {
                Emit(output, slice, top + y, page, fields);
                // Floats in headers and footers are drawn but do not push text aside.
                foreach (var f in slice.Floats)
                {
                    var item = FloatItem(f, page, section, (section.Margins.Left, section.ContentWidth), top + y);
                    if (f.BehindText) output.Items.Insert(0, item); else front.Add(item);
                }
            }
        }

        if (page.Header is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            BlockFlow.Stack(BlockFlow.Build(page.Header.Blocks, section.Margins.Left, section.ContentWidth, context), stacked);
            Story(stacked, section.HeaderDistance);
        }
        if (page.Footer is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            double h = BlockFlow.Stack(BlockFlow.Build(page.Footer.Blocks, section.Margins.Left, section.ContentWidth, context), stacked);
            Story(stacked, page.Height - section.FooterDistance - h);
        }
        foreach (var placement in page.Placed) Emit(output, placement.Slice, placement.Dx, placement.Y, page, fields);
        if (page.Notes.Count > 0)
        {
            double y = page.ContentBottom - page.NotesHeight;
            double x = section.Margins.Left;
            // A continued footnote gets a full-width separator, as in Word.
            double length = page.NotesContinued ? section.ContentWidth : Math.Min(144, section.ContentWidth / 3);
            output.Items.Add(new LineItem(x, y + NoteSeparatorHeight / 2, x + length, y + NoteSeparatorHeight / 2, 0.5, Color.Black));
            y += NoteSeparatorHeight;
            foreach (var (slice, _) in page.Notes)
            {
                Emit(output, slice, y, page, fields);
                y += slice.Height + slice.TrailingSpace;
            }
        }
        foreach (var region in page.Regions)
        {
            if (!region.Section.Columns.Separator || region.Columns.Count < 2 || region.Bottom <= region.Top) continue;
            for (int i = 0; i + 1 < region.Columns.Count; i++)
            {
                double x = (region.Columns[i].X + region.Columns[i].Width + region.Columns[i + 1].X) / 2;
                output.Items.Add(new LineItem(x, region.Top, x, region.Bottom, 0.5, Color.Black));
            }
        }
        output.Items.AddRange(front);
        return output;
    }

    /// <summary>Positions a floating image on <paramref name="page"/>; <paramref name="anchorY"/> is the top of its paragraph.</summary>
    private static ImageItem FloatItem(FloatingImage f, Page page, Section section, (double X, double Width) column, double anchorY)
    {
        var m = section.Margins;
        var (hx, hw) = f.Horizontal.RelativeTo switch
        {
            FloatReference.Page => (0, page.Width),
            FloatReference.Column or FloatReference.Character => column,
            FloatReference.LeftMargin or FloatReference.InsideMargin => (0, m.Left),
            FloatReference.RightMargin or FloatReference.OutsideMargin => (page.Width - m.Right, m.Right),
            _ => (m.Left, page.Width - m.Left - m.Right),
        };
        var (vy, vh) = f.Vertical.RelativeTo switch
        {
            FloatReference.Page => (0, page.Height),
            FloatReference.Margin => (m.Top, page.Height - m.Top - m.Bottom),
            FloatReference.TopMargin => (0, m.Top),
            FloatReference.BottomMargin => (page.Height - m.Bottom, m.Bottom),
            _ => (anchorY, 0.0),
        };
        static double Place(FloatPosition p, double start, double size, double extent) => p.Alignment switch
        {
            FloatAlignment.Start => start,
            FloatAlignment.Center => start + (size - extent) / 2,
            FloatAlignment.End => start + size - extent,
            _ => start + p.Offset,
        };
        double x = Place(f.Horizontal, hx, hw, f.Width);
        double y = Place(f.Vertical, vy, vh, f.Height);
        // Keep the image on the page, as Word does for objects anchored to text.
        y = Math.Clamp(y, 0, Math.Max(0, page.Height - f.Height));
        return new ImageItem(x, y, f.Width, f.Height, f.Image);
    }

    private static void Emit(LayoutPage output, Slice slice, double y, Page page, FieldValues fields) =>
        Emit(output, slice, 0, y, page, fields);

    private static void Emit(LayoutPage output, Slice slice, double dx, double y, Page page, FieldValues fields)
    {
        foreach (var item in slice.Items)
        {
            var moved = item.Offset(dx, y);
            if (moved is TextItem { Field: { } field } text)
                moved = text with { Text = fields.Text(field, page), Glyphs = null };
            output.Items.Add(moved);
        }
    }

    internal static string Format(int number, PageNumberFormat format) => format switch
    {
        PageNumberFormat.UpperRoman => Roman(number),
        PageNumberFormat.LowerRoman => Roman(number).ToLowerInvariant(),
        PageNumberFormat.UpperLetter => Letters(number),
        PageNumberFormat.LowerLetter => Letters(number).ToLowerInvariant(),
        _ => number.ToString(CultureInfo.InvariantCulture),
    };

    private static string Roman(int n)
    {
        if (n <= 0 || n >= 4000) return n.ToString(CultureInfo.InvariantCulture);
        ReadOnlySpan<int> values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        string[] symbols = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < values.Length; i++)
            while (n >= values[i]) { sb.Append(symbols[i]); n -= values[i]; }
        return sb.ToString();
    }

    private static string Letters(int n) =>
        n <= 0 ? n.ToString(CultureInfo.InvariantCulture) : new string((char)('A' + (n - 1) % 26), Math.Min((n - 1) / 26 + 1, 50));

    /// <summary>
    /// Places slices into the columns of pages, moving keep-chains and splitting rows as needed. Slices
    /// carry absolute x coordinates for the column they were built for; placement records the shift.
    /// </summary>
    private sealed class Paginator(List<Page> pages, Func<Page> newPage, Func<Note, List<Slice>> noteLines, List<Slice> carry)
    {
        private readonly HashSet<Note> _placedNotes = new(ReferenceEqualityComparer.Instance);

        private readonly HashSet<FloatingImage> _placedFloats = new(ReferenceEqualityComparer.Instance);

        private Page Current => pages[^1];

        /// <summary>Absolute x and width of the column that receives the next slice.</summary>
        public (double X, double Width) Column => Current.Region.Columns[Current.Column];

        /// <summary>Floats the next block must wrap around, or null when there are none on the page.</summary>
        public Obstacles? Obstacles => Current.Exclusions.Count == 0 ? null
            : new Obstacles(Current.Exclusions, Column.X, Column.Width, Current.Cursor);

        /// <summary>Positions floats anchored at the current position and makes text wrap around them.</summary>
        public void PlaceFloats(IEnumerable<FloatingImage> floats) => PlaceFloats(Current, floats, Current.Cursor);

        private void PlaceFloats(Page page, IEnumerable<FloatingImage> floats, double anchorY)
        {
            foreach (var f in floats)
            {
                if (!_placedFloats.Add(f)) continue;
                var item = FloatItem(f, page, page.Region.Section, page.Region.Columns[page.Column], anchorY);
                page.Floats.Add((item, f.BehindText));
                if (f.Wrap == TextWrap.None) continue;
                var d = f.Distance;
                page.Exclusions.Add(new Exclusion(item.X - d.Left, item.Y - d.Top, item.Width + d.Left + d.Right, item.Height + d.Top + d.Bottom,
                    FullWidth: f.Wrap == TextWrap.TopAndBottom));
            }
        }

        /// <summary>Starts a section's columns: below earlier content on the current page, or on a fresh page.</summary>
        public void StartSection(Section section)
        {
            var page = Current;
            if (ReferenceEquals(page.Region.Section, section) && page.Placed.Count == 0) return;
            if (page.Region.Bottom > page.BodyBottom - 36)
            {
                NextPage(natural: false); // no useful room left; the new page starts with this section's columns
                return;
            }
            page.StartRegion(section, Math.Max(page.Region.Bottom, page.ContentTop));
        }

        public void Add(Slice slice, double builtX)
        {
            if (slice.BreakBefore && Current.HasBody) NextPage(natural: false);
            else if (slice.ColumnBreakBefore && Current.ColumnHasBody) NextColumn(natural: false);
            Place(slice, builtX, allowMove: true);
        }

        /// <summary>True when <paramref name="slice"/> cannot start in the current, non-empty column.</summary>
        public bool StartsInNextColumn(Slice slice) =>
            Current.ColumnHasBody && !slice.BreakBefore && !slice.KeepWithPrevious
            && (slice.ColumnBreakBefore || slice.Height > Current.BodyBottom - Current.Cursor + Epsilon);

        public void NextColumn(bool natural = true)
        {
            var page = Current;
            if (page.Column + 1 >= page.Region.Columns.Count)
            {
                NextPage(natural);
                return;
            }
            page.Column++;
            page.ColumnStart = page.Placed.Count;
            page.Cursor = page.Region.Top;
            page.NaturalBreak = natural;
        }

        private void NextPage(bool natural = true)
        {
            var page = newPage();
            page.NaturalBreak = natural;
            pages.Add(page);
            if (carry.Count == 0) return;
            // Continued footnote lines come first, leaving room for some body text.
            page.NotesContinued = true;
            while (carry.Count > 0 && (page.Notes.Count == 0 || page.BodyBottom - carry[0].Height - page.ContentTop > 72))
            {
                page.Notes.Add((carry[0], null));
                carry.RemoveAt(0);
            }
        }

        /// <summary>Puts footnote lines that are still waiting on new pages (after the last body text).</summary>
        public void FlushNotes()
        {
            for (int guard = 0; carry.Count > 0 && guard < 1000; guard++) NextPage(natural: false);
        }

        /// <summary>Lines of the footnotes first referenced by <paramref name="slice"/>.</summary>
        private List<(Slice Line, Note Note)> NewNotes(Slice slice) =>
            slice.Footnotes.Where(n => !_placedNotes.Contains(n)).Distinct(ReferenceEqualityComparer.Instance).Cast<Note>()
                 .SelectMany(n => noteLines(n).Select(line => (line, n))).ToList();

        /// <summary>Lowest content of the page outside the current column (it must stay above the footnotes).</summary>
        private static double OtherColumnsBottom(Page page)
        {
            double bottom = 0;
            for (int i = 0; i < page.ColumnStart; i++)
                bottom = Math.Max(bottom, page.Placed[i].Y + page.Placed[i].Slice.Height);
            return bottom;
        }

        /// <summary>
        /// Places a slice together with its footnotes when the slice and at least the first footnote line fit
        /// (or unconditionally with <paramref name="force"/>); footnote lines that do not fit continue on the
        /// next page. Returns false when nothing was placed.
        /// </summary>
        private bool TryPutWithNotes(Page page, Slice slice, double builtX, List<(Slice Line, Note Note)> notes, bool force)
        {
            double skip = page.NaturalBreak && page.ColumnCount == 0 ? slice.LeadingSpace : 0;
            double limit = Math.Max(page.Cursor + slice.Height - skip, OtherColumnsBottom(page));
            bool Fits(Slice line) => limit <= page.BodyBottom - (page.Notes.Count == 0 ? NoteSeparatorHeight : 0) - line.Height + Epsilon;
            // Lines still waiting from earlier footnotes go first, so new ones can only follow them.
            bool fits = carry.Count > 0 ? limit <= page.BodyBottom + Epsilon : Fits(notes[0].Line);
            if (!fits && !force) return false;

            Put(page, slice, builtX);
            foreach (var (_, note) in notes) _placedNotes.Add(note);
            foreach (var (line, _) in notes)
            {
                if (carry.Count == 0 && Fits(line)) page.Notes.Add((line, slice));
                else carry.Add(line);
            }
            return true;
        }

        private void Place(Slice slice, double builtX, bool allowMove)
        {
            for (int guard = 0; guard < 10_000; guard++)
            {
                var page = Current;
                if (!page.ColumnHasBody && slice.RepeatHeader is { } header && !slice.IsHeaderRow && page.ColumnCount == 0)
                    foreach (var h in header) Put(page, h, builtX, repeated: true);

                double available = page.BodyBottom - page.Cursor;
                if (page.NaturalBreak && page.ColumnCount == 0) available += slice.LeadingSpace;
                var notes = slice.Footnotes.Count > 0 ? NewNotes(slice) : [];
                if (notes.Count > 0)
                {
                    if (TryPutWithNotes(page, slice, builtX, notes, force: false)) return;
                    if (!page.ColumnHasBody && page.Notes.Count == 0)
                    {
                        TryPutWithNotes(page, slice, builtX, notes, force: true); // nothing else can be done
                        return;
                    }
                }
                else if (slice.Height <= available + Epsilon)
                {
                    Put(page, slice, builtX);
                    return;
                }

                if (!page.ColumnHasBody && notes.Count == 0)
                {
                    // Empty column: split if possible, otherwise let it overflow (nothing else can be done).
                    if (slice.Split(available) is { } parts && parts.First.Height > Epsilon)
                    {
                        Put(page, parts.First, builtX);
                        NextColumn();
                        slice = parts.Remainder;
                        continue;
                    }
                    Put(page, slice, builtX);
                    return;
                }

                if (allowMove && slice.KeepWithPrevious && MoveChain(page, slice, builtX)) return;

                if (notes.Count == 0 && slice.Split(available) is { } split && split.First.Height > Epsilon)
                {
                    Put(page, split.First, builtX);
                    NextColumn();
                    slice = split.Remainder;
                    continue;
                }
                NextColumn();
            }
        }

        /// <summary>Moves the run of slices bound to <paramref name="slice"/> to the next column, if that helps.</summary>
        private bool MoveChain(Page page, Slice slice, double builtX)
        {
            int start = page.Placed.Count - 1;
            while (start >= page.ColumnStart && page.Placed[start].Slice.KeepWithPrevious && !page.IsRepeated(page.Placed[start].Slice)) start--;
            if (start < page.ColumnStart) return false; // chain begins at the column top: moving it would not help
            int firstBody = page.Placed.FindIndex(page.ColumnStart, p => !page.IsRepeated(p.Slice));
            if (start <= firstBody) return false;
            if (page.Placed[start].Slice.BreakBefore || page.Placed[start].Slice.ColumnBreakBefore) return false;

            var moved = page.Placed.GetRange(start, page.Placed.Count - start);
            page.Placed.RemoveRange(start, page.Placed.Count - start);
            // Their footnotes move with them.
            var owners = moved.Select(m => m.Slice).ToHashSet();
            if (page.Notes.Exists(n => n.Owner is not null && owners.Contains(n.Owner)) || moved.Any(m => m.Slice.Footnotes.Count > 0))
            {
                foreach (var note in moved.SelectMany(m => m.Slice.Footnotes))
                {
                    _placedNotes.Remove(note);
                    foreach (var line in noteLines(note)) carry.Remove(line);
                }
                page.Notes.RemoveAll(n => n.Owner is not null && owners.Contains(n.Owner));
            }
            var last = page.Placed[^1];
            page.Cursor = last.Y + last.Slice.Height + last.Slice.TrailingSpace;
            NextColumn();
            foreach (var m in moved) Place(m.Slice, BuiltX(m), allowMove: false);
            Place(slice, builtX, allowMove: false);
            return true;
        }

        private static double BuiltX(Placement p) => p.Region.Columns[p.Column].X - p.Dx;

        private void Put(Page page, Slice slice, double builtX, bool repeated = false)
        {
            double skip = page.NaturalBreak && page.ColumnCount == 0 ? slice.LeadingSpace : 0;
            var region = page.Region;
            if (slice.Floats.Count > 0) PlaceFloats(page, slice.Floats, page.Cursor - skip);
            page.Placed.Add(new Placement(slice, region.Columns[page.Column].X - builtX, page.Cursor - skip, region, page.Column));
            if (repeated) page.Repeated.Add(slice);
            page.Cursor += slice.Height + slice.TrailingSpace - skip;
            region.Bottom = Math.Max(region.Bottom, Math.Min(page.Cursor, page.BodyBottom));
        }

        /// <summary>
        /// Redistributes the current region's slices on the last page so the columns end at about the same
        /// height (smallest column height that still fits everything, found by bisection).
        /// </summary>
        public void Balance()
        {
            var page = Current;
            var region = page.Region;
            int n = region.Columns.Count;
            int first = page.Placed.FindIndex(p => ReferenceEquals(p.Region, region));
            if (n < 2 || first < 0) return;
            var items = page.Placed.GetRange(first, page.Placed.Count - first);
            if (items.Any(p => page.IsRepeated(p.Slice))) return; // keep repeated table headers where they are

            double limit = page.BodyBottom - region.Top;
            List<(int Column, double Y)>? Fill(double height)
            {
                var result = new List<(int, double)>(items.Count);
                int column = 0;
                double y = 0;
                foreach (var item in items)
                {
                    var slice = item.Slice;
                    if (y > 0 && y + slice.Height > height + Epsilon)
                    {
                        if (++column >= n) return null;
                        y = 0;
                    }
                    double skip = y == 0 && column > 0 ? slice.LeadingSpace : 0;
                    if (slice.Height - skip > height + Epsilon) return null;
                    result.Add((column, y - skip));
                    y += slice.Height + slice.TrailingSpace - skip;
                }
                return result;
            }

            if (Fill(limit) is null) return;
            double lo = 0, hi = limit;
            for (int i = 0; i < 30 && hi - lo > 0.5; i++)
            {
                double mid = (lo + hi) / 2;
                if (Fill(mid) is null) lo = mid; else hi = mid;
            }
            var layout = Fill(hi)!;
            double bottom = region.Top;
            for (int i = 0; i < items.Count; i++)
            {
                var (column, y) = layout[i];
                var item = items[i];
                double built = BuiltX(item);
                page.Placed[first + i] = item with { Column = column, Y = region.Top + y, Dx = region.Columns[column].X - built };
                bottom = Math.Max(bottom, region.Top + y + item.Slice.Height + item.Slice.TrailingSpace);
            }
            region.Bottom = Math.Min(bottom, page.BodyBottom);
            page.Column = layout[^1].Column;
            page.ColumnStart = first + layout.FindIndex(l => l.Column == page.Column);
            page.Cursor = region.Bottom;
        }
    }
}
