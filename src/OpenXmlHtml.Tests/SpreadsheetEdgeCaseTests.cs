[TestFixture]
public class SpreadsheetEdgeCaseTests
{
    [Test]
    public Task UnclosedTags() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b>bold <i>italic"));

    [Test]
    public Task ExtraClosingTags() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("text</b></i>more"));

    [Test]
    public Task ConsecutiveBreaks() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("one<br><br><br>two"));

    [Test]
    public Task WhitespaceCollapsing() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("  lots   of    spaces  "));

    [Test]
    public Task TabsAndNewlines() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("text\twith\ttabs\nand\nnewlines"));

    [Test]
    public Task SpecialCharacters() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("price: $100 & tax < 10% > 5%"));

    [Test]
    public Task UnknownTags() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<custom>text</custom>"));

    [Test]
    public Task ImageAlt() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("before <img alt=\"image description\"> after"));

    [Test]
    public Task SpanWithNoStyle() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<span>plain span</span>"));

    [Test]
    public Task MultipleSpaces() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("one     two"));

    [Test]
    public Task EmptyTags() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b></b><i></i>text"));

    [Test]
    public Task MalformedHtml() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b>bold <i>overlap</b> still italic</i>"));

    [Test]
    public Task NumericEntity() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("&#169; copyright"));

    [Test]
    public Task CiteTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<cite>citation</cite>"));

    [Test]
    public Task DfnTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<dfn>definition</dfn>"));

    [Test]
    public Task VarTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<var>variable</var>"));

    [Test]
    public Task SampTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<samp>sample output</samp>"));
}
