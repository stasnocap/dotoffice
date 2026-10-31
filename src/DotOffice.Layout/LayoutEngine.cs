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
        if (usesPageRefs) pages = Paginate(document, Bookmarks(pages), out _, cancellationToken);

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
            if (field.Kind != FieldKind.PageRef) return "1";
            pageRefs = true;
            return bookmarks is not null && field.Bookmark is not null && bookmarks.TryGetValue(field.Bookmark, out var label)
                ? Format(label.Number, field.NumberFormat ?? label.Format)
                : field.CachedResult is { Length: > 0 } cached ? cached : "1";
        });

        Page? current = null;
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

            var paginator = new Paginator(pages, () =>
            {
                var page = NewPage(section, NextNumber(section, ref numberingSection, ref pageNumber), false, settings);
                current = page;
                return page;
            });
            paginator.StartSection(section);
            bool keepWithNext = false;
            for (int i = 0; i < section.Blocks.Count; i++)
            {
                // Blocks are built for the column they start in, so columns of different widths work.
                var (x, width) = paginator.Column;
                bool keep = keepWithNext;
                var slices = BlockFlow.BuildBlock(section.Blocks, i, x, width, bodyContext, ref keepWithNext);
                if (section.Columns.Definitions.Count > 1 && slices.Count > 0 && paginator.StartsInNextColumn(slices[0]))
                {
                    paginator.NextColumn();
                    (x, width) = paginator.Column;
                    keepWithNext = keep;
                    slices = BlockFlow.BuildBlock(section.Blocks, i, x, width, bodyContext, ref keepWithNext);
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
        usesPageRefs = pageRefs;
        return pages;
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

        if (page.Header is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            BlockFlow.Stack(BlockFlow.Build(page.Header.Blocks, section.Margins.Left, section.ContentWidth, context), stacked);
            foreach (var (y, slice) in stacked) Emit(output, slice, section.HeaderDistance + y, page, fields);
        }
        if (page.Footer is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            double h = BlockFlow.Stack(BlockFlow.Build(page.Footer.Blocks, section.Margins.Left, section.ContentWidth, context), stacked);
            double top = page.Height - section.FooterDistance - h;
            foreach (var (y, slice) in stacked) Emit(output, slice, top + y, page, fields);
        }
        foreach (var placement in page.Placed) Emit(output, placement.Slice, placement.Dx, placement.Y, page, fields);
        foreach (var region in page.Regions)
        {
            if (!region.Section.Columns.Separator || region.Columns.Count < 2 || region.Bottom <= region.Top) continue;
            for (int i = 0; i + 1 < region.Columns.Count; i++)
            {
                double x = (region.Columns[i].X + region.Columns[i].Width + region.Columns[i + 1].X) / 2;
                output.Items.Add(new LineItem(x, region.Top, x, region.Bottom, 0.5, Color.Black));
            }
        }
        return output;
    }

    private static void Emit(LayoutPage output, Slice slice, double y, Page page, FieldValues fields) =>
        Emit(output, slice, 0, y, page, fields);

    private static void Emit(LayoutPage output, Slice slice, double dx, double y, Page page, FieldValues fields)
    {
        foreach (var item in slice.Items)
        {
            var moved = item.Offset(dx, y);
            if (moved is TextItem { Field: { } field } text)
                moved = text with { Text = fields.Text(field, page) };
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
    private sealed class Paginator(List<Page> pages, Func<Page> newPage)
    {
        private Page Current => pages[^1];

        /// <summary>Absolute x and width of the column that receives the next slice.</summary>
        public (double X, double Width) Column => Current.Region.Columns[Current.Column];

        /// <summary>Starts a section's columns: below earlier content on the current page, or on a fresh page.</summary>
        public void StartSection(Section section)
        {
            var page = Current;
            if (ReferenceEquals(page.Region.Section, section) && page.Placed.Count == 0) return;
            if (page.Region.Bottom > page.ContentBottom - 36)
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
            && (slice.ColumnBreakBefore || slice.Height > Current.ContentBottom - Current.Cursor + Epsilon);

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
        }

        private void Place(Slice slice, double builtX, bool allowMove)
        {
            for (int guard = 0; guard < 10_000; guard++)
            {
                var page = Current;
                if (!page.ColumnHasBody && slice.RepeatHeader is { } header && !slice.IsHeaderRow && page.ColumnCount == 0)
                    foreach (var h in header) Put(page, h, builtX, repeated: true);

                double available = page.ContentBottom - page.Cursor;
                if (page.NaturalBreak && page.ColumnCount == 0) available += slice.LeadingSpace;
                if (slice.Height <= available + Epsilon)
                {
                    Put(page, slice, builtX);
                    return;
                }

                if (!page.ColumnHasBody)
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

                if (slice.Split(available) is { } split && split.First.Height > Epsilon)
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
            var last = page.Placed[^1];
            page.Cursor = last.Y + last.Slice.Height + last.Slice.TrailingSpace;
            NextColumn();
            foreach (var m in moved) Place(m.Slice, BuiltX(m), allowMove: false);
            Place(slice, builtX, allowMove: false);
            return true;
        }

        private static double BuiltX(Placement p) => p.Region.Columns[p.Column].X - p.Dx;

        private static void Put(Page page, Slice slice, double builtX, bool repeated = false)
        {
            double skip = page.NaturalBreak && page.ColumnCount == 0 ? slice.LeadingSpace : 0;
            var region = page.Region;
            page.Placed.Add(new Placement(slice, region.Columns[page.Column].X - builtX, page.Cursor - skip, region, page.Column));
            if (repeated) page.Repeated.Add(slice);
            page.Cursor += slice.Height + slice.TrailingSpace - skip;
            region.Bottom = Math.Max(region.Bottom, Math.Min(page.Cursor, page.ContentBottom));
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

            double limit = page.ContentBottom - region.Top;
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
            region.Bottom = Math.Min(bottom, page.ContentBottom);
            page.Column = layout[^1].Column;
            page.ColumnStart = first + layout.FindIndex(l => l.Column == page.Column);
            page.Cursor = region.Bottom;
        }
    }
}
