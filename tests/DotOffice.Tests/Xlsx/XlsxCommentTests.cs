using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxCommentTests
{
    private static (Worksheet Sheet, ConversionWarningList Warnings) Read(string pageSetup, string comments)
    {
        var warnings = new ConversionWarningList();
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: pageSetup);
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}comments\" Target=\"../comments1.xml\"/></Relationships>");
        builder.AddPart("xl/comments1.xml", comments);
        return (XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings }).Sheets[0], warnings);
    }

    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    private const string Comments =
        "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>Ann</author><author>Bob</author></authors><commentList>" +
        "<comment ref=\"B3\" authorId=\"1\"><text><r><rPr><b/></rPr><t>Bob:</t></r><r><t xml:space=\"preserve\">\nsecond</t></r></text></comment>" +
        "<comment ref=\"$A$1\" authorId=\"0\"><text><t>first</t></text></comment></commentList></comments>";

    [Fact]
    public void Reads_comments_in_cell_order_with_author_and_joined_runs()
    {
        var (sheet, warnings) = Read("<pageSetup cellComments=\"atEnd\"/>", Comments);
        Assert.Equal(CommentPrinting.AtEnd, sheet.PageSetup.Comments);
        Assert.Equal([new SheetComment(0, 0, "A1", "Ann", "first"), new SheetComment(2, 1, "B3", "Bob", "Bob:\nsecond")], sheet.Comments);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Comments_are_not_printed_unless_the_page_setup_asks()
    {
        var (sheet, warnings) = Read("", Comments);
        Assert.Equal(2, sheet.Comments.Count);
        Assert.Equal(CommentPrinting.None, sheet.PageSetup.Comments);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Printing_comments_as_displayed_falls_back_to_the_list_with_a_warning()
    {
        var (sheet, warnings) = Read("<pageSetup cellComments=\"asDisplayed\"/>", Comments);
        Assert.Equal(CommentPrinting.AsDisplayed, sheet.PageSetup.Comments);
        Assert.Equal(["xlsx.comments"], warnings.Warnings.Select(w => w.Code));
    }

    [Fact]
    public void A_malformed_comments_part_warns()
    {
        var (sheet, warnings) = Read("<pageSetup cellComments=\"atEnd\"/>", "<comments><oops>");
        Assert.Empty(sheet.Comments);
        Assert.Equal(["xlsx.comments"], warnings.Warnings.Select(w => w.Code));
    }
}
