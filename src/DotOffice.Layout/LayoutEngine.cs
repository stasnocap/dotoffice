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
        public List<(Slice Slice, double Y)> Placed { get; } = [];

        /// <summary>The page was started because content overflowed (not by a forced break or section start).</summary>
        public bool NaturalBreak { get; set; }
        public double Cursor { get; set; }
        public bool HasBody => Placed.Exists(p => !p.Slice.IsHeaderRow || !IsRepeated(p.Slice));
        public HashSet<Slice> Repeated { get; } = [];
        public bool IsRepeated(Slice s) => Repeated.Contains(s);
    }

    public LayoutDocument Layout(Document document, CancellationToken cancellationToken = default)
    {
        var pages = new List<Page>();
        var settings = document.Settings;
        var bodyContext = Context(settings, kind => kind == FieldKind.Page ? "1" : "1");

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

            var slices = BlockFlow.Build(section.Blocks, section.Margins.Left, section.ContentWidth, bodyContext);
            var paginator = new Paginator(pages, () =>
            {
                var page = NewPage(section, NextNumber(section, ref numberingSection, ref pageNumber), false, settings);
                current = page;
                return page;
            });
            foreach (var slice in slices) paginator.Add(slice);
            current = pages[^1];
        }

        var result = new LayoutDocument { Title = document.Title, Author = document.Author };
        int total = pages.Count;
        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            result.Pages.Add(Render(page, total, settings));
        }
        return result;
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

    private FlowContext Context(DocumentSettings settings, Func<FieldKind, string> fieldText) => new()
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
        var placeholder = Context(settings, kind => kind == FieldKind.Page ? Format(number, section.PageNumberFormat) : "99");
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

        return new Page
        {
            Section = section, Width = section.PageWidth, Height = section.PageHeight,
            ContentTop = top, ContentBottom = bottom, Cursor = top,
            Header = header, Footer = footer, Number = number,
        };
    }

    private LayoutPage Render(Page page, int total, DocumentSettings settings)
    {
        var output = new LayoutPage(page.Width, page.Height);
        var section = page.Section;
        string pageText = Format(page.Number, section.PageNumberFormat);
        string totalText = total.ToString(CultureInfo.InvariantCulture);
        var context = Context(settings, kind => kind == FieldKind.Page ? pageText : totalText);

        if (page.Header is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            BlockFlow.Stack(BlockFlow.Build(page.Header.Blocks, section.Margins.Left, section.ContentWidth, context), stacked);
            foreach (var (y, slice) in stacked) Emit(output, slice, section.HeaderDistance + y, pageText, totalText);
        }
        if (page.Footer is not null)
        {
            var stacked = new List<(double Y, Slice Slice)>();
            double h = BlockFlow.Stack(BlockFlow.Build(page.Footer.Blocks, section.Margins.Left, section.ContentWidth, context), stacked);
            double top = page.Height - section.FooterDistance - h;
            foreach (var (y, slice) in stacked) Emit(output, slice, top + y, pageText, totalText);
        }
        foreach (var (slice, y) in page.Placed) Emit(output, slice, y, pageText, totalText);
        return output;
    }

    private static void Emit(LayoutPage page, Slice slice, double y, string pageText, string totalText)
    {
        foreach (var item in slice.Items)
        {
            var moved = item.Offset(0, y);
            if (moved is TextItem { Field: { } field } text)
                moved = text with { Text = field == FieldKind.Page ? pageText : totalText };
            page.Items.Add(moved);
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

    /// <summary>Places slices onto pages, moving keep-chains and splitting rows as needed.</summary>
    private sealed class Paginator(List<Page> pages, Func<Page> newPage)
    {
        private Page Current => pages[^1];

        public void Add(Slice slice)
        {
            if (slice.BreakBefore && Current.HasBody) NextPage(natural: false);
            Place(slice, allowMove: true);
        }

        private void NextPage(bool natural = true)
        {
            var page = newPage();
            page.NaturalBreak = natural;
            pages.Add(page);
        }

        private void Place(Slice slice, bool allowMove)
        {
            for (int guard = 0; guard < 10_000; guard++)
            {
                var page = Current;
                if (!page.HasBody && slice.RepeatHeader is { } header && !slice.IsHeaderRow && page.Placed.Count == 0)
                    foreach (var h in header) Put(page, h, repeated: true);

                double available = page.ContentBottom - page.Cursor;
                if (page.NaturalBreak && page.Placed.Count == 0) available += slice.LeadingSpace;
                if (slice.Height <= available + Epsilon)
                {
                    Put(page, slice);
                    return;
                }

                if (!page.HasBody)
                {
                    // Empty page: split if possible, otherwise let it overflow (nothing else can be done).
                    if (slice.Split(available) is { } parts && parts.First.Height > Epsilon)
                    {
                        Put(page, parts.First);
                        NextPage();
                        slice = parts.Remainder;
                        continue;
                    }
                    Put(page, slice);
                    return;
                }

                if (allowMove && slice.KeepWithPrevious && MoveChain(page, slice)) return;

                if (slice.Split(available) is { } split && split.First.Height > Epsilon)
                {
                    Put(page, split.First);
                    NextPage();
                    slice = split.Remainder;
                    continue;
                }
                NextPage();
            }
        }

        /// <summary>Moves the run of slices bound to <paramref name="slice"/> to a new page, if that helps.</summary>
        private bool MoveChain(Page page, Slice slice)
        {
            int start = page.Placed.Count - 1;
            while (start >= 0 && page.Placed[start].Slice.KeepWithPrevious && !page.IsRepeated(page.Placed[start].Slice)) start--;
            if (start < 0) return false; // chain begins at the page top: moving it would not help
            int firstBody = page.Placed.FindIndex(p => !page.IsRepeated(p.Slice));
            if (start <= firstBody) return false;
            if (page.Placed[start].Slice.BreakBefore) return false;

            var moved = page.Placed.GetRange(start, page.Placed.Count - start).Select(p => p.Slice).ToList();
            page.Placed.RemoveRange(start, page.Placed.Count - start);
            page.Cursor = page.Placed.Count == 0 ? page.ContentTop : page.Placed[^1].Y + page.Placed[^1].Slice.Height + page.Placed[^1].Slice.TrailingSpace;
            NextPage();
            foreach (var s in moved) Place(s, allowMove: false);
            Place(slice, allowMove: false);
            return true;
        }

        private static void Put(Page page, Slice slice, bool repeated = false)
        {
            double skip = page.NaturalBreak && page.Placed.Count == 0 ? slice.LeadingSpace : 0;
            page.Placed.Add((slice, page.Cursor - skip));
            if (repeated) page.Repeated.Add(slice);
            page.Cursor += slice.Height + slice.TrailingSpace - skip;
        }
    }
}