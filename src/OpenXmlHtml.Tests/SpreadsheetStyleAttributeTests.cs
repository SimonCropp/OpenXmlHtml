[TestFixture]
public class SpreadsheetStyleAttributeTests
{
    [Test]
    public Task FontWeightBold() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-weight: bold\">bold</span>"));

    [Test]
    public Task FontWeight700() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-weight: 700\">bold</span>"));

    [Test]
    public Task FontStyleItalic() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-style: italic\">italic</span>"));

    [Test]
    public Task TextDecorationUnderline() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"text-decoration: underline\">underlined</span>"));

    [Test]
    public Task TextDecorationLineThrough() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"text-decoration: line-through\">struck</span>"));

    [Test]
    public Task MultipleStyleProperties() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-weight: bold; font-style: italic; color: #FF0000\">styled</span>"));

    [Test]
    public Task VerticalAlignSuper() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "E = mc<span style=\"vertical-align: super\">2</span>"));

    [Test]
    public Task VerticalAlignSub() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "H<span style=\"vertical-align: sub\">2</span>O"));
}
