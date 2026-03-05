[TestFixture]
public class SpreadsheetColorTests
{
    [Test]
    public Task FontColorAttribute() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<font color=\"#FF0000\">red text</font>"));

    [Test]
    public Task FontColorShortHex() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<font color=\"#F00\">red text</font>"));

    [Test]
    public Task NamedColor() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<span style=\"color: blue\">blue text</span>"));

    [Test]
    public Task RgbColor() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<span style=\"color: rgb(0, 128, 0)\">green text</span>"));

    [Test]
    public Task MultipleColors() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            """
            <span style="color: red">red</span>
            <span style="color: blue">blue</span>
            <span style="color: green">green</span>
            """));

    [Test]
    public Task ColorWithFormatting() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<b style=\"color: #FF0000\">bold red</b>"));

    [Test]
    public Task NestedColors() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"color: red\">outer <span style=\"color: blue\">inner</span> outer</span>"));
}
