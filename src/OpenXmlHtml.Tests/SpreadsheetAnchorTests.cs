[TestFixture]
public class SpreadsheetAnchorTests
{
    [Test]
    public Task SimpleLink() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<a href=\"https://example.com\">Example</a>"));

    [Test]
    public Task LinkWithSameText() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<a href=\"https://example.com\">https://example.com</a>"));

    [Test]
    public Task LinkWithFormatting() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<a href=\"https://example.com\"><b>Bold Link</b></a>"));

    [Test]
    public Task LinkInText() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("Visit <a href=\"https://example.com\">our site</a> for more info."));

    [Test]
    public Task LinkWithNoHref() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<a>anchor text</a>"));

    [Test]
    public Task MultipleLinks() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<a href=\"https://one.com\">One</a> and <a href=\"https://two.com\">Two</a>"));
}
