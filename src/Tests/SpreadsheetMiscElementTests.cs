[TestFixture]
public class SpreadsheetMiscElementTests
{
    [Test]
    public Task AbbrTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "The <abbr title=\"World Health Organization\">WHO</abbr> recommends it."));

    [Test]
    public Task AcronymTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "Use <acronym title=\"HyperText Markup Language\">HTML</acronym> for web pages."));

    [Test]
    public Task TimeTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "The meeting is at <time datetime=\"14:00\">2 PM</time>."));

    [Test]
    public Task QTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "She said <q>hello world</q> to everyone."));

    [Test]
    public Task NestedQ() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<q>outer <q>inner</q> outer</q>"));

    [Test]
    public Task FigcaptionTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<figure><img alt=\"Chart\"><figcaption>Figure 1: Sales data</figcaption></figure>"));

    [Test]
    public Task SvgTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "before <svg width=\"100\" height=\"100\"><circle cx=\"50\" cy=\"50\" r=\"40\"/></svg> after"));

    [Test]
    public Task ArticleTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<article>Article content here</article>"));

    [Test]
    public Task AsideTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<aside>Sidebar content</aside>"));

    [Test]
    public Task SectionTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<section>Section content</section>"));

    [Test]
    public Task DtWithBold() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<dl><dt>Term</dt><dd>Definition of the term</dd></dl>"));

    [Test]
    public Task BlockquoteWithQ() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<blockquote><q>To be or not to be</q></blockquote>"));
}
