using DotOffice.Converter;

namespace DotOffice.Tests;

public class DocumentConverterTests
{
    [Fact]
    public async Task Unsupported_conversion_throws()
    {
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            DocumentConverter.ConvertAsync(Stream.Null, Stream.Null, DocumentFormat.Docx, DocumentFormat.Pdf));
    }
}
