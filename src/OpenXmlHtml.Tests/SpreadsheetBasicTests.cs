[TestFixture]
public class SpreadsheetBasicTests
{
    [Test]
    public Task PlainText() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("Hello world"));

    [Test]
    public Task Bold() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b>bold text</b>"));

    [Test]
    public Task Strong() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<strong>strong text</strong>"));

    [Test]
    public Task Italic() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<i>italic text</i>"));

    [Test]
    public Task Em() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<em>emphasized</em>"));

    [Test]
    public Task Underline() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<u>underlined</u>"));

    [Test]
    public Task Strikethrough() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<s>struck</s>"));

    [Test]
    public Task StrikeTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<strike>struck</strike>"));

    [Test]
    public Task Del() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<del>deleted</del>"));

    [Test]
    public Task Superscript() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("x<sup>2</sup>"));

    [Test]
    public Task Subscript() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("H<sub>2</sub>O"));

    [Test]
    public Task LineBreak() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("line one<br>line two"));

    [Test]
    public Task SelfClosingBreak() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("line one<br/>line two"));

    [Test]
    public Task MixedFormatting() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("normal <b>bold</b> <i>italic</i> normal"));

    [Test]
    public Task EmptyHtml() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(""));

    [Test]
    public Task WhitespaceOnly() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("   "));

    [Test]
    public Task HtmlEntities() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("&amp; &lt; &gt; &quot; &apos;"));

    [Test]
    public Task NonBreakingSpace() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("hello&nbsp;world"));

    [Test]
    public Task InsTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ins>inserted</ins>"));
}
