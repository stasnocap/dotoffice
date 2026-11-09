using DotOffice.Core;
using DotOffice.Docx;
using static DotOffice.Tests.Docx.DocxBuilder;

namespace DotOffice.Tests.Docx;

/// <summary>Word 2016's defaults for missing w:docDefaults children (measured, differs from ISO 29500); see StyleSheet.</summary>
public class DocDefaultsTests
{
    private const string RPr = "<w:rPrDefault><w:rPr/></w:rPrDefault>";
    private const string PPr = "<w:pPrDefault><w:pPr/></w:pPrDefault>";

    private static (RunProperties Run, ParagraphProperties Para) Read(string? styles, string runXml = "", bool explicitDefaults = false)
    {
        var doc = DocxReader.Read(new DocxBuilder { Styles = styles, ExplicitDefaults = explicitDefaults, Body = P(R("x", runXml)) }.BuildStream(), new DocxReaderOptions());
        var p = doc.Sections.SelectMany(s => s.Blocks).OfType<Paragraph>().First();
        return (p.Inlines.OfType<TextRun>().First().Properties, p.Properties);
    }

    private static string Defaults(string inner) => $"<w:docDefaults>{inner}</w:docDefaults>";

    private static void AssertWord(RunProperties r) => Assert.Equal(("Calibri", 11.0), (r.FontFamily, r.FontSize));
    private static void AssertSpec(RunProperties r) => Assert.Equal(("Times New Roman", 10.0), (r.FontFamily, r.FontSize));

    private static void AssertWordPara(ParagraphProperties p)
    {
        Assert.Equal(8, p.SpaceAfter);
        Assert.Equal(259 / 240.0, p.LineSpacing.Value, 6);
        Assert.Equal(LineSpacingRule.Auto, p.LineSpacing.Rule);
    }

    private static void AssertSpecPara(ParagraphProperties p)
    {
        Assert.Equal(0, p.SpaceAfter);
        Assert.Equal(LineSpacing.Single, p.LineSpacing);
    }

    [Fact]
    public void Absent_rPrDefault_is_Calibri_11()
    {
        AssertWord(Read(Defaults(PPr)).Run);
        AssertWord(Read(Defaults("")).Run);
        AssertWord(Read("<w:docDefaults/>").Run);
    }

    [Fact]
    public void Present_empty_rPrDefault_keeps_spec_defaults()
    {
        AssertSpec(Read(Defaults(RPr)).Run);
        AssertSpec(Read(Defaults("<w:rPrDefault/>")).Run);
    }

    [Fact]
    public void Present_rPrDefault_stating_font_only_has_spec_size()
    {
        var r = Read(Defaults("<w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr></w:rPrDefault>")).Run;
        Assert.Equal(("Arial", 10.0), (r.FontFamily, r.FontSize));
    }

    [Fact]
    public void Absent_pPrDefault_is_1_08_and_8_pt_after()
    {
        AssertWordPara(Read(Defaults(RPr)).Para);
        AssertWordPara(Read(Defaults("")).Para);
    }

    [Fact]
    public void Bare_or_empty_pPrDefault_keeps_spec_defaults()
    {
        AssertSpecPara(Read(Defaults("<w:pPrDefault/>")).Para);
        AssertSpecPara(Read(Defaults(PPr)).Para);
    }

    [Fact]
    public void Present_pPrDefault_stating_jc_only_has_spec_spacing()
    {
        var p = Read(Defaults("<w:pPrDefault><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:pPrDefault>")).Para;
        AssertSpecPara(p);
        Assert.Equal(ParagraphAlignment.Center, p.Alignment);
    }

    [Fact]
    public void Present_pPrDefault_stating_after_only_or_line_only_merges_per_attribute()
    {
        var after = Read(Defaults("<w:pPrDefault><w:pPr><w:spacing w:after=\"100\"/></w:pPr></w:pPrDefault>")).Para;
        Assert.Equal((5.0, LineSpacing.Single), (after.SpaceAfter, after.LineSpacing));
        var line = Read(Defaults("<w:pPrDefault><w:pPr><w:spacing w:line=\"276\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault>")).Para;
        Assert.Equal(0, line.SpaceAfter);
        Assert.Equal(276 / 240.0, line.LineSpacing.Value, 6);
    }

    [Fact]
    public void The_two_elements_are_independent()
    {
        var (r1, p1) = Read(Defaults(RPr));
        AssertSpec(r1); AssertWordPara(p1);
        var (r2, p2) = Read(Defaults(PPr));
        AssertWord(r2); AssertSpecPara(p2);
    }

    [Fact]
    public void No_styles_part_gives_Word_defaults()
    {
        var (r, p) = Read(null);
        AssertWord(r); AssertWordPara(p);
    }

    [Fact]
    public void Styles_without_Normal_style_or_docDefaults_give_Word_defaults()
    {
        var (r, p) = Read("<w:style w:type=\"paragraph\" w:styleId=\"Other\"><w:name w:val=\"Other\"/></w:style>");
        AssertWord(r); AssertWordPara(p);
    }

    [Fact]
    public void Test_helper_states_empty_defaults_unless_asked_not_to()
    {
        var (r, p) = Read(null, explicitDefaults: true);
        AssertSpec(r); AssertSpecPara(p);
    }

    private const string Normal = "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/>{0}</w:style>";

    [Fact]
    public void Normal_stating_spacing_overrides_Word_defaults()
    {
        var (r, p) = Read(string.Format(Normal, "<w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>"));
        AssertWord(r); AssertSpecPara(p);
    }

    [Fact]
    public void Normal_stating_only_after_keeps_1_08_line()
    {
        var (_, p) = Read(string.Format(Normal, "<w:pPr><w:spacing w:after=\"0\"/></w:pPr>"));
        Assert.Equal(0, p.SpaceAfter);
        Assert.Equal(259 / 240.0, p.LineSpacing.Value, 6);
    }

    [Fact]
    public void Normal_or_run_stating_Arial_12_keeps_1_08_and_8_pt_after()
    {
        var (r, p) = Read(string.Format(Normal, "<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"24\"/></w:rPr>"));
        Assert.Equal(("Arial", 12.0), (r.FontFamily, r.FontSize));
        AssertWordPara(p);

        var (r2, p2) = Read(null, "<w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"24\"/>");
        Assert.Equal(("Arial", 12.0), (r2.FontFamily, r2.FontSize));
        AssertWordPara(p2);
    }

    [Fact]
    public void Bare_pPrDefault_with_Arial_12_rPrDefault_is_single_with_no_gaps()
    {
        var (r, p) = Read(Defaults("<w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"24\"/></w:rPr></w:rPrDefault><w:pPrDefault/>"));
        Assert.Equal(("Arial", 12.0), (r.FontFamily, r.FontSize));
        AssertSpecPara(p);
    }
}
