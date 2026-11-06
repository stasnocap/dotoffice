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

    /// <summary>A footnote line waiting to be placed, built for the note area at <see cref="X"/> of <see cref="Width"/>.</summary>
    private readonly record struct NoteLine(Slice Line, Note Note, double X, double Width);

    /// <summary>Footnote lines at the bottom of one column (the whole text width in single-column text).</summary>
    private sealed class NoteArea(double x, double width)
    {
        public double X { get; } = x;
        public double Width { get; } = width;

        /// <summary>Footnote lines with the body slice that referenced them (null when carried over).</summary>
        public List<(Slice Slice, Slice? Owner)> Lines { get; } = [];

        /// <summary>The first lines continue a footnote from the previous page.</summary>
        public bool Continued { get; set; }

        public double Height => Lines.Count == 0 ? 0 : NoteSeparatorHeight + Lines.Sum(n => n.Slice.Height + n.Slice.TrailingSpace);

        public bool Overlaps((double X, double Width) column) => X < column.X + column.Width - Epsilon && column.X < X + Width - Epsilon;
    }

    private sealed class Page
    {
        public required Section Section { get; init; }

        /// <summary>The section's margins, left and right swapped on even pages with mirrored margins.</summary>
        public required Thickness Margins { get; init; }
        public bool Mirrored { get; init; }

        /// <summary>Margins of <paramref name="section"/> as they apply on this page.</summary>
        public Thickness MarginsOf(Section section) =>
            Mirrored ? new Thickness(section.Margins.Top, section.Margins.Left, section.Margins.Bottom, section.Margins.Right) : section.Margins;
        public double Width { get; init; }
        public double Height { get; init; }
        public bool IsSectionStart { get; init; }
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

        /// <summary>Footnote areas at the page bottom, one per column that has footnotes.</summary>
        public List<NoteArea> NoteAreas { get; } = [];

        /// <summary>
        /// The footnote area for <paramref name="column"/>: one already holding lines below it (when the page
        /// mixes column layouts) or else one of the column's own width.
        /// </summary>
        public NoteArea NoteAreaFor((double X, double Width) column)
        {
            var area = NoteAreas.Find(a => a.Lines.Count > 0 && a.Overlaps(column))
                       ?? NoteAreas.Find(a => Math.Abs(a.X - column.X) < 0.5 && Math.Abs(a.Width - column.Width) < 0.5);
            if (area is null) NoteAreas.Add(area = new NoteArea(column.X, column.Width));
            return area;
        }

        /// <summary>Bottom of the body text in <paramref name="column"/>: the content area minus the footnotes below it.</summary>
        public double BodyBottomAt((double X, double Width) column) =>
            ContentBottom - NoteAreas.Where(a => a.Overlaps(column)).Select(a => a.Height).DefaultIfEmpty(0).Max();

        /// <summary>Bottom of the body text in the current column.</summary>
        public double BodyBottom => BodyBottomAt(Region.Columns[Column]);

        /// <summary>Bottom of the body text in the column with the most footnotes.</summary>
        public double LowestBodyBottom => ContentBottom - NoteAreas.Select(a => a.Height).DefaultIfEmpty(0).Max();

        /// <summary>Floating images, painted below (<c>Behind</c>) or above the text.</summary>
        public List<PlacedFloat> Floats { get; } = [];

        /// <summary>Areas wrapped text avoids.</summary>
        public List<Exclusion> Exclusions { get; } = [];

        public void StartRegion(Section section, double top)
        {
            var columns = section.Columns.Resolve(section.ContentWidth).Select(c => (MarginsOf(section).Left + c.X, c.Width)).ToList();
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
        var lineNumbers = NumberLines(pages);
        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            result.Pages.Add(Render(page, fields, settings, lineNumbers));
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
        var carry = new List<NoteLine>(); // footnote lines continued on the next page
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

            var notes = new Dictionary<(Note, double, double), List<Slice>>();
            var paginator = new Paginator(pages, () =>
            {
                var page = NewPage(section, NextNumber(section, ref numberingSection, ref pageNumber), false, settings);
                current = page;
                return page;
            }, (note, x, width) =>
            {
                if (!notes.TryGetValue((note, x, width), out var lines))
                    notes[(note, x, width)] = lines = BlockFlow.Build(note.Blocks, x, width, bodyContext);
                return lines;
            }, carry);
            last = paginator;
            paginator.StartSection(section);
            bool keepWithNext = false;
            for (int i = 0; i < section.Blocks.Count; i++)
            {
                if (section.Blocks[i] is Paragraph { Properties.Frame: { } frame })
                {
                    // Consecutive paragraphs with equal frames share one frame.
                    var group = new List<Paragraph>();
                    int end = i;
                    while (end < section.Blocks.Count && section.Blocks[end] is Paragraph fp && fp.Properties.Frame == frame) group.Add((Paragraph)section.Blocks[end++]);
                    var (fx, fw) = paginator.Column;
                    var box = FloatingTable.BuildFrame(group, frame, fw, paginator.BodyHeight, bodyContext);
                    if (box is not null)
                    {
                        paginator.PlaceTable(box);
                        i = end - 1;
                        continue;
                    }
                    bodyContext.Warnings.Warn("docx.frame", "A paragraph frame that does not fit on a page is laid out in the text flow.");
                }
                if (section.Blocks[i] is Table { Properties.Position: not null } floating)
                {
                    var (fx, fw) = paginator.Column;
                    var box = FloatingTable.Build(floating, fw, paginator.BodyHeight, bodyContext);
                    if (box is not null)
                    {
                        paginator.PlaceTable(box);
                        continue;
                    }
                    bodyContext.Warnings.Warn("docx.floating-table", "A floating table that does not fit on a page is laid out in the text flow.");
                }
                // Blocks are built for the column they start in, so columns of different widths work.
                var (x, width) = paginator.Column;
                bool keep = keepWithNext;
                var floats = section.Blocks[i] is Paragraph p ? p.Inlines.OfType<FloatingObject>().Where(f => !f.Properties.Hidden).ToList() : [];
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
                for (int j = 0, rebuilds = 0; j < slices.Count; j++)
                {
                    // A paragraph continuing in a column of another width is broken again for that width.
                    if (slices[j] is { Reflow: { } reflow, BreakBefore: false, ColumnBreakBefore: false } && paginator.StartsInNextColumn(slices[j]))
                    {
                        paginator.NextColumn();
                        if (Math.Abs(paginator.Column.Width - width) > 0.5)
                        {
                            (x, width) = paginator.Column;
                            slices = [.. slices.Take(j), .. reflow(x, width, paginator.Obstacles)];
                        }
                    }
                    paginator.Add(slices[j], x);

                    // Lines moved on into a column of another width (keep-with-next, widow control): lay them out again there.
                    if (Math.Abs(paginator.Column.Width - width) < 0.5 || rebuilds >= 4) continue;
                    int k = paginator.TrailingInColumn(slices, j);
                    if (k < 0 || (k > 0 && slices[k].Reflow is null)) continue;
                    paginator.Unplace(slices.GetRange(k, j - k + 1));
                    rebuilds++;
                    (x, width) = paginator.Column;
                    if (k == 0)
                    {
                        keepWithNext = keep;
                        slices = BlockFlow.BuildBlock(section.Blocks, i, x, width, bodyContext, ref keepWithNext, paginator.Obstacles);
                    }
                    else slices = [.. slices.Take(k), .. slices[k].Reflow!(x, width, paginator.Obstacles)];
                    j = k - 1;
                }
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

        bool mirrored = settings.MirrorMargins && number % 2 == 0;
        var margins = mirrored
            ? new Thickness(section.Margins.Top, section.Margins.Left, section.Margins.Bottom, section.Margins.Right)
            : section.Margins;
        double top = section.Margins.Top, bottom = section.PageHeight - section.Margins.Bottom;
        var placeholder = Context(settings, field => field.Kind == FieldKind.Page
            ? Format(number, field.NumberFormat ?? section.PageNumberFormat)
            : field.Kind == FieldKind.PageRef ? field.CachedResult ?? "99" : "99");
        if (header is not null)
        {
            double h = BlockFlow.Stack(BlockFlow.Build(header.Blocks, margins.Left, section.ContentWidth, placeholder), []);
            top = Math.Max(top, section.HeaderDistance + h);
        }
        if (footer is not null)
        {
            double h = BlockFlow.Stack(BlockFlow.Build(footer.Blocks, margins.Left, section.ContentWidth, placeholder), []);
            bottom = Math.Min(bottom, section.PageHeight - section.FooterDistance - h);
        }
        if (bottom - top < 36) bottom = top + 36; // absurd margins: keep a usable body area

        var page = new Page
        {
            Section = section, Margins = margins, Mirrored = mirrored, Width = section.PageWidth, Height = section.PageHeight,
            ContentTop = top, ContentBottom = bottom,
            Header = header, Footer = footer, Number = number, IsSectionStart = sectionStart,
        };
        page.StartRegion(section, top);
        return page;
    }

    /// <summary>The line number of each numbered paragraph line (w:lnNumType): restarted per page or section, or running on.</summary>
    private static Dictionary<Slice, int> NumberLines(List<Page> pages)
    {
        var numbers = new Dictionary<Slice, int>(ReferenceEqualityComparer.Instance);
        int count = 0;
        Section? previous = null;
        foreach (var page in pages)
        {
            bool newPage = true;
            foreach (var placement in page.Placed)
            {
                if (placement.Slice.LineNumber is null || placement.Region.Section.LineNumbering is not { } rule) continue;
                if (!ReferenceEquals(previous, placement.Region.Section) && (rule.Restart == LineNumberRestart.NewSection || previous is null)
                    || newPage && rule.Restart == LineNumberRestart.NewPage)
                    count = rule.Start - 1;
                previous = placement.Region.Section;
                newPage = false;
                numbers[placement.Slice] = ++count;
            }
        }
        return numbers;
    }

    /// <summary>Word's automatic distance between line numbers and text (a guess: 14.5 pt, as LibreOffice).</summary>
    private const double AutoLineNumberDistance = 14.5;

    private LayoutPage Render(Page page, FieldValues fields, DocumentSettings settings, Dictionary<Slice, int> lineNumbers)
    {
        var output = new LayoutPage(page.Width, page.Height);
        var section = page.Section;
        var context = Context(settings, field => fields.Text(field, page));
        var front = new List<LayoutItem>();
        foreach (var placed in page.Floats)
        {
            var items = Realize(placed.Object, placed.X, placed.Y, context);
            if (placed.Behind) output.Items.AddRange(items); else front.AddRange(items);
        }

        void Story(List<(double Y, Slice Slice)> stacked, double top)
        {
            foreach (var (y, slice) in stacked)
            {
                Emit(output, slice, top + y, page, fields);
                // Floats in headers and footers are drawn but do not push text aside.
                foreach (var f in slice.Floats)
                {
                    var (fx, fy) = FloatOrigin(f, page, section, (page.Margins.Left, section.ContentWidth), top + y);
                    var items = Realize(f, fx, fy, context);
                    if (f.BehindText) output.Items.InsertRange(0, items); else front.AddRange(items);
                }
            }
        }

        if (page.Header is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            BlockFlow.Stack(BlockFlow.Build(page.Header.Blocks, page.Margins.Left, section.ContentWidth, context), stacked);
            Story(stacked, section.HeaderDistance);
        }
        if (page.Footer is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            double h = BlockFlow.Stack(BlockFlow.Build(page.Footer.Blocks, page.Margins.Left, section.ContentWidth, context), stacked);
            Story(stacked, page.Height - section.FooterDistance - h);
        }
        foreach (var placement in page.Placed)
        {
            Emit(output, placement.Slice, placement.Dx, placement.Y, page, fields);
            if (!lineNumbers.TryGetValue(placement.Slice, out int number) || placement.Slice.LineNumber is not { } style) continue;
            var rule = placement.Region.Section.LineNumbering!;
            if (number % rule.CountBy != 0) continue;
            string text = number.ToString(CultureInfo.InvariantCulture);
            double right = placement.Region.Columns[placement.Column].X - (rule.Distance ?? AutoLineNumberDistance);
            output.Items.Add(new TextItem(right - style.Font.MeasureText(text, style.Size), placement.Y + style.Baseline, style.Font, style.Size, text, Color.Black));
        }
        foreach (var area in page.NoteAreas)
        {
            if (area.Lines.Count == 0) continue;
            double y = page.ContentBottom - area.Height;
            double x = area.X;
            // A continued footnote gets a full-width separator, as in Word.
            double length = area.Continued ? area.Width : Math.Min(144, area.Width / 3);
            output.Items.Add(new LineItem(x, y + NoteSeparatorHeight / 2, x + length, y + NoteSeparatorHeight / 2, 0.5, Color.Black));
            y += NoteSeparatorHeight;
            foreach (var (slice, _) in area.Lines)
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
        if (section.PageBorders is { } pageBorders && ShowsPageBorders(pageBorders, page.IsSectionStart))
        {
            var items = new List<LayoutItem>();
            AddPageBorders(items, pageBorders, section, page.Margins);
            if (pageBorders.InFront) output.Items.AddRange(items); else output.Items.InsertRange(0, items);
        }
        return output;
    }

    private static bool ShowsPageBorders(PageBorders borders, bool firstOfSection) => borders.Display switch
    {
        PageBorderDisplay.FirstPage => firstOfSection,
        PageBorderDisplay.NotFirstPage => !firstOfSection,
        _ => true,
    };

    /// <summary>
    /// Draws w:pgBorders. The space is measured to the nearest edge of the line: from the page edge outward-in, or from the
    /// text margin outward. Top and bottom lines run over the full outer width so the corners meet.
    /// </summary>
    internal static void AddPageBorders(List<LayoutItem> items, PageBorders b, Section section, Thickness m)
    {
        double W = section.PageWidth, H = section.PageHeight;
        // Centre line of each edge.
        double top = b.OffsetFromText ? m.Top - b.Top.Space - b.Top.Width / 2 : b.Top.Space + b.Top.Width / 2;
        double bottom = b.OffsetFromText ? H - m.Bottom + b.Bottom.Space + b.Bottom.Width / 2 : H - b.Bottom.Space - b.Bottom.Width / 2;
        double left = b.OffsetFromText ? m.Left - b.Left.Space - b.Left.Width / 2 : b.Left.Space + b.Left.Width / 2;
        double right = b.OffsetFromText ? W - m.Right + b.Right.Space + b.Right.Width / 2 : W - b.Right.Space - b.Right.Width / 2;
        double x1 = b.Left.IsVisible ? left - b.Left.Width / 2 : left, x2 = b.Right.IsVisible ? right + b.Right.Width / 2 : right;
        if (b.Top.IsVisible) ParagraphLayouter.AddBorder(items, b.Top, x1, top, x2, top);
        if (b.Bottom.IsVisible) ParagraphLayouter.AddBorder(items, b.Bottom, x1, bottom, x2, bottom);
        if (b.Left.IsVisible) ParagraphLayouter.AddBorder(items, b.Left, left, top, left, bottom);
        if (b.Right.IsVisible) ParagraphLayouter.AddBorder(items, b.Right, right, top, right, bottom);
    }

    /// <summary>Positions a floating image on <paramref name="page"/>; <paramref name="anchorY"/> is the top of its paragraph.</summary>
    private static (double X, double Y) FloatOrigin(FloatingObject f, Page page, Section section, (double X, double Width) column, double anchorY)
    {
        var m = page.Margins;
        var (hx, hw) = f.Horizontal.RelativeTo switch
        {
            FloatReference.Page => (0, page.Width),
            FloatReference.Column or FloatReference.Character => column,
            FloatReference.LeftMargin => (0, m.Left),
            FloatReference.RightMargin => (page.Width - m.Right, m.Right),
            // Inside is the left margin, except on the even pages of mirrored margins, where it is the right one.
            FloatReference.InsideMargin => page.Mirrored ? (page.Width - m.Right, m.Right) : (0, m.Left),
            FloatReference.OutsideMargin => page.Mirrored ? (0, m.Left) : (page.Width - m.Right, m.Right),
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
        static double Place(FloatPosition p, double start, double size, double extent, bool swap) => p.Alignment switch
        {
            FloatAlignment.Start => start,
            FloatAlignment.Center => start + (size - extent) / 2,
            FloatAlignment.End => start + size - extent,
            FloatAlignment.Inside => swap ? start + size - extent : start,
            FloatAlignment.Outside => swap ? start : start + size - extent,
            _ => start + p.Offset,
        };
        double x = Place(f.Horizontal, hx, hw, f.Width, page.Mirrored);
        double y = Place(f.Vertical, vy, vh, f.Height, false);
        // Keep the image on the page, as Word does for objects anchored to text.
        y = Math.Clamp(y, 0, Math.Max(0, page.Height - f.Height));
        return (x, y);
    }

    /// <summary>The layout items of a floating object whose top-left corner is at (<paramref name="x"/>, <paramref name="y"/>).</summary>
    private static List<LayoutItem> Realize(FloatingObject f, double x, double y, FlowContext context)
    {
        if (f is FloatingTable table) return table.Items.Select(i => i.Offset(x, y)).ToList();
        if (f is FloatingImage image) return [new ImageItem(x, y, f.Width, f.Height, image.Image)];
        return f is FloatingTextBox box ? TextBoxes.Realize(box, x, y, context) : [];
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
    private sealed class Paginator(List<Page> pages, Func<Page> newPage, Func<Note, double, double, List<Slice>> noteLines, List<NoteLine> carry)
    {
        private readonly HashSet<Note> _placedNotes = new(ReferenceEqualityComparer.Instance);

        /// <summary>Where each floating image went, so it can follow its anchor when that moves to another page.</summary>
        private readonly Dictionary<FloatingObject, (Page Page, PlacedFloat Item, Exclusion? Exclusion)> _placedFloats =
            new(ReferenceEqualityComparer.Instance);

        private Page Current => pages[^1];

        /// <summary>Absolute x and width of the column that receives the next slice.</summary>
        public (double X, double Width) Column => Current.Region.Columns[Current.Column];

        /// <summary>Floats the next block must wrap around, or null when there are none on the page.</summary>
        public Obstacles? Obstacles => Current.Exclusions.Count == 0 ? null
            : new Obstacles(Current.Exclusions, Column.X, Column.Width, Current.Cursor);

        /// <summary>Height of the body area of the current page.</summary>
        public double BodyHeight => Current.ContentBottom - Current.ContentTop;

        /// <summary>Places a floating table; one anchored to the text moves to a new page when it would not fit below the cursor.</summary>
        public void PlaceTable(FloatingTable table)
        {
            if (table.Vertical.RelativeTo == FloatReference.Paragraph && Current.HasBody
                && Current.Cursor + table.Height + table.Distance.Top + table.Distance.Bottom > Current.BodyBottom + Epsilon)
                NextPage(natural: false);
            PlaceFloats([table]);
        }

        /// <summary>Positions floats anchored at the current position and makes text wrap around them.</summary>
        public void PlaceFloats(IEnumerable<FloatingObject> floats) => PlaceFloats(Current, floats, Current.Cursor);

        private void PlaceFloats(Page page, IEnumerable<FloatingObject> floats, double anchorY)
        {
            foreach (var f in floats)
            {
                if (_placedFloats.TryGetValue(f, out var placed))
                {
                    if (ReferenceEquals(placed.Page, page)) continue;
                    // The anchor moved to a later page: the image goes with it.
                    placed.Page.Floats.Remove(placed.Item);
                    if (placed.Exclusion is { } old) placed.Page.Exclusions.Remove(old);
                }
                var (fx, fy) = FloatOrigin(f, page, page.Region.Section, page.Region.Columns[page.Column], anchorY);
                var item = new PlacedFloat(f, fx, fy, f.BehindText);
                page.Floats.Add(item);
                Exclusion? exclusion = null;
                if (f.Wrap != TextWrap.None)
                {
                    var d = f.Distance;
                    exclusion = new Exclusion(item.X - d.Left, item.Y - d.Top, f.Width + d.Left + d.Right, f.Height + d.Top + d.Bottom,
                        FullWidth: f.Wrap == TextWrap.TopAndBottom);
                    page.Exclusions.Add(exclusion.Value);
                }
                _placedFloats[f] = (page, item, exclusion);
            }
        }

        /// <summary>Starts a section's columns: below earlier content on the current page, or on a fresh page.</summary>
        public void StartSection(Section section)
        {
            var page = Current;
            if (ReferenceEquals(page.Region.Section, section) && page.Placed.Count == 0) return;
            if (page.Region.Bottom > page.LowestBodyBottom - 36)
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
            // Continued footnote lines come first, leaving room for some body text.
            while (carry.Count > 0)
            {
                var (line, _, x, width) = carry[0];
                var area = page.NoteAreaFor((x, width));
                if (area.Lines.Count > 0 && page.BodyBottomAt((x, width)) - line.Height - page.ContentTop <= 72) break;
                if (area.Lines.Count == 0) area.Continued = true;
                area.Lines.Add((line, null));
                carry.RemoveAt(0);
            }
        }

        /// <summary>Puts footnote lines that are still waiting on new pages (after the last body text).</summary>
        public void FlushNotes()
        {
            for (int guard = 0; carry.Count > 0 && guard < 1000; guard++) NextPage(natural: false);
        }

        /// <summary>Lines of the footnotes first referenced by <paramref name="slice"/>, built for <paramref name="area"/>.</summary>
        private List<NoteLine> NewNotes(Slice slice, NoteArea area) =>
            slice.Footnotes.Where(n => !_placedNotes.Contains(n)).Distinct(ReferenceEqualityComparer.Instance).Cast<Note>()
                 .SelectMany(n => noteLines(n, area.X, area.Width).Select(line => new NoteLine(line, n, area.X, area.Width))).ToList();

        /// <summary>Lowest content outside the current column above <paramref name="area"/> (it must stay above the footnotes).</summary>
        private static double OtherColumnsBottom(Page page, NoteArea area)
        {
            double bottom = 0;
            for (int i = 0; i < page.ColumnStart; i++)
            {
                var p = page.Placed[i];
                if (area.Overlaps(p.Region.Columns[p.Column])) bottom = Math.Max(bottom, p.Y + p.Slice.Height);
            }
            return bottom;
        }

        /// <summary>
        /// Places a slice together with its footnotes when the slice and at least the first footnote line fit
        /// (or unconditionally with <paramref name="force"/>); footnote lines that do not fit continue on the
        /// next page. Returns false when nothing was placed.
        /// </summary>
        private bool TryPutWithNotes(Page page, Slice slice, double builtX, NoteArea area, List<NoteLine> notes, bool force)
        {
            double skip = page.NaturalBreak && page.ColumnCount == 0 ? slice.LeadingSpace : 0;
            double limit = Math.Max(page.Cursor + slice.Height - skip, OtherColumnsBottom(page, area));
            bool Fits(Slice line) => limit <= page.ContentBottom - area.Height - (area.Lines.Count == 0 ? NoteSeparatorHeight : 0) - line.Height + Epsilon;
            // Lines still waiting from earlier footnotes go first, so new ones can only follow them.
            bool fits = carry.Count > 0 ? limit <= page.BodyBottom + Epsilon : Fits(notes[0].Line);
            if (!fits && !force) return false;

            Put(page, slice, builtX);
            foreach (var line in notes) _placedNotes.Add(line.Note);
            foreach (var line in notes)
            {
                if (carry.Count == 0 && Fits(line.Line)) area.Lines.Add((line.Line, slice));
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

                if (slice is RowSlice) SkipFloats(page, slice);
                double available = page.BodyBottom - page.Cursor;
                if (page.NaturalBreak && page.ColumnCount == 0) available += slice.LeadingSpace;
                var area = slice.Footnotes.Count > 0 ? page.NoteAreaFor(page.Region.Columns[page.Column]) : null;
                var notes = area is not null ? NewNotes(slice, area) : [];
                if (notes.Count > 0)
                {
                    if (TryPutWithNotes(page, slice, builtX, area!, notes, force: false)) return;
                    if (!page.ColumnHasBody && area!.Lines.Count == 0)
                    {
                        TryPutWithNotes(page, slice, builtX, area, notes, force: true); // nothing else can be done
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

        /// <summary>Moves the cursor below wrapped floats beside <paramref name="slice"/>: table rows do not wrap, as in Word.</summary>
        private static void SkipFloats(Page page, Slice slice)
        {
            var (x, width) = page.Region.Columns[page.Column];
            for (bool moved = true; moved;)
            {
                moved = false;
                foreach (var e in page.Exclusions)
                    if (e.X < x + width && e.Right > x && e.Y < page.Cursor + slice.Height && e.Bottom > page.Cursor)
                    {
                        page.Cursor = e.Bottom;
                        moved = true;
                    }
            }
        }

        /// <summary>
        /// The index of the first of <paramref name="slices"/>[0..<paramref name="last"/>] that, with all after it,
        /// ends the current column; -1 when the current column does not end with them.
        /// </summary>
        public int TrailingInColumn(List<Slice> slices, int last)
        {
            var page = Current;
            int k = last, p = page.Placed.Count - 1;
            if (p < page.ColumnStart || !ReferenceEquals(page.Placed[p].Slice, slices[last])) return -1;
            while (k > 0 && p > page.ColumnStart && ReferenceEquals(page.Placed[p - 1].Slice, slices[k - 1])) { k--; p--; }
            return k;
        }

        /// <summary>Removes <paramref name="slices"/>, the last placements of the current column, with their footnotes.</summary>
        public void Unplace(List<Slice> slices)
        {
            var page = Current;
            int start = page.Placed.Count - slices.Count;
            page.Placed.RemoveRange(start, slices.Count);
            ForgetNotes(page, slices);
            if (start > page.ColumnStart)
            {
                var last = page.Placed[^1];
                page.Cursor = last.Y + last.Slice.Height + last.Slice.TrailingSpace;
            }
            else page.Cursor = page.Region.Top;
            var region = page.Region;
            region.Bottom = page.Placed.Where(p => ReferenceEquals(p.Region, region))
                                .Select(p => p.Y + p.Slice.Height + p.Slice.TrailingSpace).DefaultIfEmpty(region.Top).Max();
        }

        /// <summary>Takes back the footnotes of <paramref name="slices"/> so they are placed again with them.</summary>
        private void ForgetNotes(Page page, IReadOnlyCollection<Slice> slices)
        {
            if (!slices.Any(s => s.Footnotes.Count > 0)) return;
            var owners = slices.ToHashSet();
            var notes = slices.SelectMany(s => s.Footnotes).ToHashSet<Note>(ReferenceEqualityComparer.Instance);
            foreach (var note in notes) _placedNotes.Remove(note);
            carry.RemoveAll(line => notes.Contains(line.Note));
            foreach (var area in page.NoteAreas) area.Lines.RemoveAll(n => n.Owner is not null && owners.Contains(n.Owner));
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
            ForgetNotes(page, moved.Select(m => m.Slice).ToList());
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

            double limit = region.Columns.Min(page.BodyBottomAt) - region.Top;
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
            region.Bottom = Math.Min(bottom, region.Columns.Min(page.BodyBottomAt));
            page.Column = layout[^1].Column;
            page.ColumnStart = first + layout.FindIndex(l => l.Column == page.Column);
            page.Cursor = region.Bottom;
        }
    }
}
