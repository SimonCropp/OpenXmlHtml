[TestFixture]
public class SpreadsheetNestedTests
{
    [Test]
    public Task BoldItalic() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b><i>bold italic</i></b>"));

    [Test]
    public Task BoldUnderlineItalic() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b><u><i>all three</i></u></b>"));

    [Test]
    public Task NestedSameTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b>outer <b>inner</b> outer</b>"));

    [Test]
    public Task PartialOverlap() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b>bold <i>bold-italic</i> bold</b>"));

    [Test]
    public Task DeeplyNested() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<b><i><u><s>all formats</s></u></i></b>"));

    [Test]
    public Task MixedContent() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("start <b>bold <i>both</i></b> <u>under</u> end"));
}
