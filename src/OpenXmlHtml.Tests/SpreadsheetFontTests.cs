[TestFixture]
public class SpreadsheetFontTests
{
    [Test]
    public Task FontFaceAttribute() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<font face=\"Arial\">arial text</font>"));

    [Test]
    public Task FontSizeAttribute() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<font size=\"14\">large text</font>"));

    [Test]
    public Task FontAllAttributes() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<font color=\"#0000FF\" size=\"16\" face=\"Verdana\">styled text</font>"));

    [Test]
    public Task InlineStyleFontFamily() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-family: 'Comic Sans MS'\">comic text</span>"));

    [Test]
    public Task InlineStyleFontSizePt() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-size: 18pt\">large text</span>"));

    [Test]
    public Task InlineStyleFontSizePx() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-size: 16px\">pixel sized text</span>"));

    [Test]
    public Task InlineStyleFontSizeEm() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-size: 1.5em\">em sized text</span>"));

    [Test]
    public Task InlineStyleFontSizeKeyword() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<span style=\"font-size: x-large\">x-large text</span>"));

    [Test]
    public Task SmallTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("normal <small>smaller</small> normal"));

    [Test]
    public Task CodeTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("normal <code>monospace</code> normal"));

    [Test]
    public Task KbdTag() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("Press <kbd>Ctrl+C</kbd> to copy"));

    [Test]
    public Task Base64ImageSkipped()
    {
        var png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";
        return Verify(SpreadsheetHtmlConverter.ToInlineString(
            $"""before <img src="data:image/png;base64,{png}"> after"""));
    }

    [Test]
    public Task ImageWithAltTextInSpreadsheet() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            """before <img src="https://example.com/logo.png" alt="Logo"> after"""));
}
