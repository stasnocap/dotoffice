namespace DotOffice.Core;

/// <summary>
/// Unit conventions for the Core model: every length is expressed in <b>points</b> (1/72 inch)
/// stored as <see cref="double"/>. Readers convert their native units (twips, EMU, half-points…)
/// once at load time; layout and rendering never see anything else.
/// </summary>
public static class Units
{
    public const double PointsPerInch = 72;

    /// <summary>Twentieths of a point (WordprocessingML "dxa").</summary>
    public static double FromTwips(double twips) => twips / 20.0;

    /// <summary>English Metric Units (DrawingML): 914400 per inch.</summary>
    public static double FromEmu(long emu) => emu / 12700.0;

    public static double FromHalfPoints(double halfPoints) => halfPoints / 2.0;

    public static double FromEighthPoints(double eighthPoints) => eighthPoints / 8.0;

    public static double FromInches(double inches) => inches * PointsPerInch;

    public static double FromMillimeters(double mm) => mm * PointsPerInch / 25.4;
}

/// <summary>Four edge lengths in points, e.g. page margins or cell padding.</summary>
public readonly record struct Thickness(double Top, double Right, double Bottom, double Left)
{
    public static Thickness Zero { get; } = new(0, 0, 0, 0);
}