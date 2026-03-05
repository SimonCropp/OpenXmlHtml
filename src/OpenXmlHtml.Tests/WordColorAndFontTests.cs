[TestFixture]
public class WordColorAndFontTests
{
    [Test]
    public Task FontColorAttribute() =>
        Verify(WordHtmlConverter.ToParagraphs("<font color=\"#FF0000\">red text</font>"));

    [Test]
    public Task NamedColor() =>
        Verify(WordHtmlConverter.ToParagraphs("<span style=\"color: blue\">blue text</span>"));

    [Test]
    public Task RgbColor() =>
        Verify(WordHtmlConverter.ToParagraphs("<span style=\"color: rgb(0, 128, 0)\">green text</span>"));

    [Test]
    public Task FontFace() =>
        Verify(WordHtmlConverter.ToParagraphs("<font face=\"Arial\">arial text</font>"));

    [Test]
    public Task FontSize() =>
        Verify(WordHtmlConverter.ToParagraphs("<font size=\"14\">large text</font>"));

    [Test]
    public Task InlineStyleFontFamily() =>
        Verify(WordHtmlConverter.ToParagraphs("<span style=\"font-family: Verdana\">verdana</span>"));

    [Test]
    public Task InlineStyleFontSizePt() =>
        Verify(WordHtmlConverter.ToParagraphs("<span style=\"font-size: 24pt\">big text</span>"));

    [Test]
    public Task MultipleStyles() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<span style=\"font-weight: bold; font-style: italic; color: #0000FF; font-size: 16pt\">styled</span>"));

    [Test]
    public Task CodeTag() =>
        Verify(WordHtmlConverter.ToParagraphs("Use <code>Console.WriteLine</code> to print"));

    [Test]
    public Task SmallTag() =>
        Verify(WordHtmlConverter.ToParagraphs("normal <small>smaller</small> normal"));

    [Test]
    public Task ColorWithBold() =>
        Verify(WordHtmlConverter.ToParagraphs("<b style=\"color: red\">bold red</b>"));
}
