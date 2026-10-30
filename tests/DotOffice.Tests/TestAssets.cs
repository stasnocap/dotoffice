namespace DotOffice.Tests;

internal static class TestAssets
{
    public static string Path(string name) => System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", name);

    public static byte[] Bytes(string name) => File.ReadAllBytes(Path(name));
}
