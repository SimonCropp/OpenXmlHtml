[TestFixture]
public class WordBasicTests
{
    [Test]
    public Task PlainText() =>
        Verify(WordHtmlConverter.ToParagraphs("Hello world"));

    [Test]
    public Task Bold() =>
        Verify(WordHtmlConverter.ToParagraphs("<b>bold text</b>"));

    [Test]
    public Task Strong() =>
        Verify(WordHtmlConverter.ToParagraphs("<strong>strong text</strong>"));

    [Test]
    public Task Italic() =>
        Verify(WordHtmlConverter.ToParagraphs("<i>italic text</i>"));

    [Test]
    public Task Em() =>
        Verify(WordHtmlConverter.ToParagraphs("<em>emphasized</em>"));

    [Test]
    public Task Underline() =>
        Verify(WordHtmlConverter.ToParagraphs("<u>underlined</u>"));

    [Test]
    public Task Strikethrough() =>
        Verify(WordHtmlConverter.ToParagraphs("<s>struck</s>"));

    [Test]
    public Task Del() =>
        Verify(WordHtmlConverter.ToParagraphs("<del>deleted</del>"));

    [Test]
    public Task Superscript() =>
        Verify(WordHtmlConverter.ToParagraphs("x<sup>2</sup>"));

    [Test]
    public Task Subscript() =>
        Verify(WordHtmlConverter.ToParagraphs("H<sub>2</sub>O"));

    [Test]
    public Task LineBreak() =>
        Verify(WordHtmlConverter.ToParagraphs("line one<br>line two"));

    [Test]
    public Task MixedFormatting() =>
        Verify(WordHtmlConverter.ToParagraphs("normal <b>bold</b> <i>italic</i> normal"));

    [Test]
    public Task EmptyHtml() =>
        Verify(WordHtmlConverter.ToParagraphs(""));

    [Test]
    public Task HtmlEntities() =>
        Verify(WordHtmlConverter.ToParagraphs("&amp; &lt; &gt; &quot;"));

    [Test]
    public Task InsTag() =>
        Verify(WordHtmlConverter.ToParagraphs("<ins>inserted</ins>"));
}
