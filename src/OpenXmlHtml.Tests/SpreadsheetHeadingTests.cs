[TestFixture]
public class SpreadsheetHeadingTests
{
    [Test]
    public Task H1() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h1>Main Title</h1>"));

    [Test]
    public Task H2() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h2>Subtitle</h2>"));

    [Test]
    public Task H3() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h3>Section</h3>"));

    [Test]
    public Task H4() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h4>Subsection</h4>"));

    [Test]
    public Task H5() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h5>Minor</h5>"));

    [Test]
    public Task H6() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h6>Smallest</h6>"));

    [Test]
    public Task HeadingWithInlineFormatting() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h1>Title with <i>italic</i> word</h1>"));

    [Test]
    public Task HeadingFollowedByParagraph() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h2>Heading</h2><p>Body text</p>"));
}
