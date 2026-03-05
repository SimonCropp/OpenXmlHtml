[TestFixture]
public class WordMiscElementTests
{
    [Test]
    public Task AbbrTag() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "The <abbr title=\"World Health Organization\">WHO</abbr> recommends it."));

    [Test]
    public Task AcronymTag() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "Use <acronym title=\"HyperText Markup Language\">HTML</acronym> for web pages."));

    [Test]
    public Task TimeTag() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "The meeting is at <time datetime=\"14:00\">2 PM</time>."));

    [Test]
    public Task QTag() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "She said <q>hello world</q> to everyone."));

    [Test]
    public Task FigcaptionTag() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<figure><img alt=\"Chart\"><figcaption>Figure 1: Sales data</figcaption></figure>"));

    [Test]
    public Task SvgTag() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "before <svg width=\"100\" height=\"100\"><circle cx=\"50\" cy=\"50\" r=\"40\"/></svg> after"));

    [Test]
    public Task ArticleTag() =>
        Verify(WordHtmlConverter.ToParagraphs("<article>Article content here</article>"));

    [Test]
    public Task SectionTag() =>
        Verify(WordHtmlConverter.ToParagraphs("<section>Section content</section>"));

    [Test]
    public Task DtWithBold() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<dl><dt>Term</dt><dd>Definition of the term</dd></dl>"));

    [Test]
    public Task BlockquoteWithQ() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<blockquote><q>To be or not to be</q></blockquote>"));
}
