namespace DotOffice.Fonts;

/// <summary>A source of fonts looked up by family name and style.</summary>
public interface IFontProvider
{
    /// <summary>
    /// Returns the face of <paramref name="family"/> closest to the requested style (it may not match
    /// exactly — callers check <see cref="TrueTypeFont.IsBold"/>/<see cref="TrueTypeFont.IsItalic"/>),
    /// or null when the family is unknown.
    /// </summary>
    TrueTypeFont? FindFont(string family, bool bold, bool italic);

    /// <summary>Families worth trying, in order, for characters missing from the requested font.</summary>
    IEnumerable<string> GetFallbackFamilies() => [];
}

/// <summary>Picks the face closest to a requested weight/slant, CSS-style.</summary>
internal static class FaceMatcher
{
    public static T? Best<T>(IEnumerable<T> faces, bool bold, bool italic, Func<T, (int Weight, int Width, bool Italic)> describe)
        where T : class
    {
        T? best = null;
        int bestScore = int.MaxValue;
        foreach (var face in faces)
        {
            var (weight, width, isItalic) = describe(face);
            int score = (isItalic == italic ? 0 : 10_000)
                        + Math.Abs(weight - (bold ? 700 : 400)) * 4 + (weight >= 600 == bold ? 0 : 2_000)
                        + Math.Abs(width - 5) * 300;
            if (score < bestScore) (best, bestScore) = (face, score);
        }
        return best;
    }

    public static bool Matches(TrueTypeFont font, bool bold, bool italic) =>
        IsBold(font) == bold && font.IsItalic == italic;

    public static bool IsBold(TrueTypeFont font) => font.WeightClass >= 600 || font.IsBold;
}

/// <summary>Fonts registered explicitly by the application; consulted before system fonts.</summary>
public sealed class FontCollection : IFontProvider
{
    private readonly Dictionary<string, List<TrueTypeFont>> _families = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _order = [];

    public int Count => _families.Values.Sum(f => f.Count);

    public FontCollection Add(TrueTypeFont font)
    {
        lock (_families)
        {
            Register(font.FamilyName, font);
            if (!string.Equals(font.TypographicFamilyName, font.FamilyName, StringComparison.OrdinalIgnoreCase))
                Register(font.TypographicFamilyName, font);
        }
        return this;
    }

    public FontCollection Add(byte[] fontData) => Add(TrueTypeFont.Load(fontData));

    public FontCollection AddFile(string path) => Add(TrueTypeFont.Load(path));

    /// <summary>Registers every supported .ttf font in a directory tree, skipping unreadable files.</summary>
    public FontCollection AddDirectory(string path)
    {
        foreach (var file in FontDirectories.EnumerateFontFiles([path]))
        {
            try { AddFile(file); }
            catch (Exception e) when (e is InvalidDataException or NotSupportedException or IOException
                                          or ArgumentOutOfRangeException or IndexOutOfRangeException) { }
        }
        return this;
    }

    public TrueTypeFont? FindFont(string family, bool bold, bool italic)
    {
        lock (_families)
            return _families.TryGetValue(family, out var faces)
                ? FaceMatcher.Best(faces, bold, italic, f => (f.WeightClass, 5, f.IsItalic))
                : null;
    }

    public IEnumerable<string> GetFallbackFamilies()
    {
        lock (_families) return _order.ToArray();
    }

    private void Register(string family, TrueTypeFont font)
    {
        if (!_families.TryGetValue(family, out var list))
        {
            _families[family] = list = [];
            _order.Add(family);
        }
        if (!list.Contains(font)) list.Add(font);
    }
}

/// <summary>
/// Fonts installed on the machine, indexed lazily from the platform's standard font directories.
/// Only headers are read while indexing; a face is fully parsed the first time it is requested.
/// </summary>
public sealed class SystemFontProvider : IFontProvider
{
    private static readonly Lazy<SystemFontProvider> SharedInstance = new(() => new SystemFontProvider(FontDirectories.Default()));

    private readonly Lazy<Dictionary<string, List<FontFaceInfo>>> _index;
    private readonly Dictionary<string, TrueTypeFont?> _loaded = [];

    public SystemFontProvider(IEnumerable<string> directories)
    {
        var dirs = directories.ToArray();
        _index = new Lazy<Dictionary<string, List<FontFaceInfo>>>(() => BuildIndex(dirs));
    }

    /// <summary>Process-wide instance over the default directories (indexed once).</summary>
    public static SystemFontProvider Shared => SharedInstance.Value;

    public IReadOnlyCollection<string> Families => _index.Value.Keys;

    public TrueTypeFont? FindFont(string family, bool bold, bool italic)
    {
        if (!_index.Value.TryGetValue(family, out var faces)) return null;
        var info = FaceMatcher.Best(faces, bold, italic, f => (f.Weight, f.WidthClass, f.IsItalic));
        return info is null ? null : Load(info.Path);
    }

    private TrueTypeFont? Load(string path)
    {
        lock (_loaded)
        {
            if (_loaded.TryGetValue(path, out var font)) return font;
            try { font = TrueTypeFont.Load(path); }
            catch (Exception e) when (e is InvalidDataException or NotSupportedException or IOException
                                          or ArgumentOutOfRangeException or IndexOutOfRangeException) { font = null; }
            _loaded[path] = font;
            return font;
        }
    }

    private static Dictionary<string, List<FontFaceInfo>> BuildIndex(string[] directories)
    {
        var index = new Dictionary<string, List<FontFaceInfo>>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in FontDirectories.EnumerateFontFiles(directories))
        {
            if (FontFaceInfo.TryRead(file) is not { } info) continue;
            Add(info.FamilyName, info);
            if (!string.Equals(info.TypographicFamilyName, info.FamilyName, StringComparison.OrdinalIgnoreCase))
                Add(info.TypographicFamilyName, info);
        }
        return index;

        void Add(string family, FontFaceInfo info)
        {
            if (!index.TryGetValue(family, out var list)) index[family] = list = [];
            list.Add(info);
        }
    }
}

public static class FontDirectories
{
    /// <summary>Standard font directories for the current OS (only existing ones are returned).</summary>
    public static IReadOnlyList<string> Default()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var dirs = new List<string>();
        if (OperatingSystem.IsWindows())
        {
            dirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (local.Length > 0) dirs.Add(Path.Combine(local, "Microsoft", "Windows", "Fonts"));
        }
        else if (OperatingSystem.IsMacOS())
        {
            dirs.AddRange(["/System/Library/Fonts", "/Library/Fonts"]);
            if (home.Length > 0) dirs.Add(Path.Combine(home, "Library", "Fonts"));
        }
        else
        {
            var dataHome = Environment.GetEnvironmentVariable("XDG_DATA_HOME");
            if (string.IsNullOrEmpty(dataHome) && home.Length > 0) dataHome = Path.Combine(home, ".local", "share");
            dirs.AddRange(["/usr/share/fonts", "/usr/local/share/fonts"]);
            if (!string.IsNullOrEmpty(dataHome)) dirs.Add(Path.Combine(dataHome, "fonts"));
            if (home.Length > 0) dirs.Add(Path.Combine(home, ".fonts"));
        }
        return dirs.Where(d => d.Length > 0 && Directory.Exists(d)).Distinct().ToList();
    }

    internal static IEnumerable<string> EnumerateFontFiles(IEnumerable<string> directories)
    {
        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true, MatchCasing = MatchCasing.CaseInsensitive };
        foreach (var dir in directories)
        {
            if (!Directory.Exists(dir)) continue;
            foreach (var file in Directory.EnumerateFiles(dir, "*.ttf", options)) yield return file;
        }
    }
}