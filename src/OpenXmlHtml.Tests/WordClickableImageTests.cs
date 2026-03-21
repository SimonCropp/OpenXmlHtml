[TestFixture]
public class WordClickableImageTests
{
    [Test]
    public Task ImageInsideLink()
    {
        var png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            $"""<p><a href="https://example.com"><img src="data:image/png;base64,{png}" width="50" height="50"></a></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ImageAndTextInsideLink()
    {
        var png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            $"""<p><a href="https://example.com"><img src="data:image/png;base64,{png}" width="30" height="30"> Visit site</a></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ExternalLinkCreatesHyperlink()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Click <a href="https://example.com">here</a> for details.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ExternalLinkFallbackWithoutMainPart()
    {
        var elements = WordHtmlConverter.ToElements(
            """<p>Click <a href="https://example.com">here</a> for details.</p>""");
        return Verify(elements);
    }
}
