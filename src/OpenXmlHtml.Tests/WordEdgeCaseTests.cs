[TestFixture]
public class WordEdgeCaseTests
{
    [Test]
    public Task UnclosedTags() =>
        Verify(WordHtmlConverter.ToParagraphs("<b>bold <i>italic"));

    [Test]
    public Task ConsecutiveBreaks() =>
        Verify(WordHtmlConverter.ToParagraphs("one<br><br><br>two"));

    [Test]
    public Task WhitespaceCollapsing() =>
        Verify(WordHtmlConverter.ToParagraphs("  lots   of    spaces  "));

    [Test]
    public Task SpecialCharacters() =>
        Verify(WordHtmlConverter.ToParagraphs("price: $100 &amp; tax &lt; 10%"));

    [Test]
    public Task UnknownTags() =>
        Verify(WordHtmlConverter.ToParagraphs("<custom>text</custom>"));

    [Test]
    public Task ImageAlt() =>
        Verify(WordHtmlConverter.ToParagraphs("before <img alt=\"image description\"> after"));

    [Test]
    public Task EmptyTags() =>
        Verify(WordHtmlConverter.ToParagraphs("<b></b><i></i>text"));

    [Test]
    public Task MalformedHtml() =>
        Verify(WordHtmlConverter.ToParagraphs("<b>bold <i>overlap</b> still italic</i>"));

    [Test]
    public Task CiteTag() =>
        Verify(WordHtmlConverter.ToParagraphs("<cite>citation</cite>"));

    [Test]
    public Task VarTag() =>
        Verify(WordHtmlConverter.ToParagraphs("<var>variable</var>"));
}
