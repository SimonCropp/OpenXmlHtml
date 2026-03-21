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

    [Test]
    public Task InternalAnchorLink() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <p><a href="#section2">Jump to Section 2</a></p>
            <h2 id="section2">Section 2</h2>
            <p>Content here.</p>
            """));

    [Test]
    public Task BookmarkOnElement() =>
        Verify(WordHtmlConverter.ToElements(
            """<h1 id="intro">Introduction</h1><p>Some text.</p>"""));

    [Test]
    public Task BookmarksAndLinksDocx()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <h1>Table of Contents</h1>
            <p><a href="#chapter1">Chapter 1: Getting Started</a></p>
            <p><a href="#chapter2">Chapter 2: Advanced Topics</a></p>
            <h1 id="chapter1" style="page-break-before: always">Chapter 1: Getting Started</h1>
            <p>Welcome to the guide.</p>
            <h1 id="chapter2" style="page-break-before: always">Chapter 2: Advanced Topics</h1>
            <p>Deep dive into the subject.</p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
