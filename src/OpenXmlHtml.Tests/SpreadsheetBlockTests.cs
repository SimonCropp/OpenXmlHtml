[TestFixture]
public class SpreadsheetBlockTests
{
    [Test]
    public Task Paragraphs() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<p>first paragraph</p><p>second paragraph</p>"));

    [Test]
    public Task Divs() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<div>first div</div><div>second div</div>"));

    [Test]
    public Task Headings() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<h1>heading one</h1><h2>heading two</h2><h3>heading three</h3>"));

    [Test]
    public Task MixedBlocksAndInline() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<p>paragraph with <b>bold</b></p><div>div with <i>italic</i></div>"));

    [Test]
    public Task Blockquote() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<blockquote>quoted text</blockquote>"));

    [Test]
    public Task PreformattedText() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<pre>  line one\n  line two</pre>"));

    [Test]
    public Task HorizontalRule() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("above<hr>below"));

    [Test]
    public Task DefinitionList() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<dl><dt>Term</dt><dd>Definition</dd></dl>"));
}
