[TestFixture]
public class WordPageSizeTests
{
    [Test]
    public void ConvertToDocxEmitsA4PageSize()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx("<p>Body</p>", stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        var pageSize = document.MainDocumentPart!
            .Document!.Body!
            .GetFirstChild<SectionProperties>()!
            .GetFirstChild<PageSize>()!;

        Assert.That(pageSize.Width!.Value, Is.EqualTo(11906u));
        Assert.That(pageSize.Height!.Value, Is.EqualTo(16838u));
    }

    [Test]
    public void SetHeaderEmitsA4PageSize()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            main.Document = new(body);

            WordHtmlConverter.AppendHtml(body, "<p>Body</p>", main);
            WordHtmlConverter.SetHeader(main, "<p>Header</p>");
        }

        stream.Position = 0;
        using var opened = WordprocessingDocument.Open(stream, false);
        var pageSize = opened.MainDocumentPart!
            .Document!.Body!
            .GetFirstChild<SectionProperties>()!
            .GetFirstChild<PageSize>()!;

        Assert.That(pageSize.Width!.Value, Is.EqualTo(11906u));
        Assert.That(pageSize.Height!.Value, Is.EqualTo(16838u));
    }
}
