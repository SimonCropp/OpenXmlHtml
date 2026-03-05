[TestFixture]
public class WordAnchorTests
{
    [Test]
    public Task SimpleLink() =>
        Verify(WordHtmlConverter.ToParagraphs("<a href=\"https://example.com\">Example</a>"));

    [Test]
    public Task LinkWithSameText() =>
        Verify(WordHtmlConverter.ToParagraphs("<a href=\"https://example.com\">https://example.com</a>"));

    [Test]
    public Task LinkWithFormatting() =>
        Verify(WordHtmlConverter.ToParagraphs("<a href=\"https://example.com\"><b>Bold Link</b></a>"));

    [Test]
    public Task LinkInText() =>
        Verify(WordHtmlConverter.ToParagraphs("Visit <a href=\"https://example.com\">our site</a> for more info."));

    [Test]
    public Task LinkWithNoHref() =>
        Verify(WordHtmlConverter.ToParagraphs("<a>anchor text</a>"));
}
