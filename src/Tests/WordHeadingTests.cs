[TestFixture]
public class WordHeadingTests
{
    [Test]
    public Task H1() =>
        Verify(WordHtmlConverter.ToParagraphs("<h1>Main Title</h1>"));

    [Test]
    public Task H2() =>
        Verify(WordHtmlConverter.ToParagraphs("<h2>Subtitle</h2>"));

    [Test]
    public Task H3() =>
        Verify(WordHtmlConverter.ToParagraphs("<h3>Section</h3>"));

    [Test]
    public Task HeadingWithInlineFormatting() =>
        Verify(WordHtmlConverter.ToParagraphs("<h1>Title with <i>italic</i> word</h1>"));

    [Test]
    public Task HeadingFollowedByParagraph() =>
        Verify(WordHtmlConverter.ToParagraphs("<h2>Heading</h2><p>Body text</p>"));
}
