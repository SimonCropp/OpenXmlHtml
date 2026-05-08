[TestFixture]
public class WordLetterSpacingTests
{
    [Test]
    public Task LetterSpacingPx() =>
        Verify(WordHtmlConverter.ToParagraphs(
            """<p><span style="letter-spacing: 4px">spaced out</span></p>"""));

    [Test]
    public Task LetterSpacingPt() =>
        Verify(WordHtmlConverter.ToParagraphs(
            """<p><span style="letter-spacing: 2pt">two point spacing</span></p>"""));

    [Test]
    public Task LetterSpacingNormal() =>
        Verify(WordHtmlConverter.ToParagraphs(
            """<p><span style="letter-spacing: normal">default spacing</span></p>"""));

    [Test]
    public Task LetterSpacingConvertToDocx()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="letter-spacing: 3px">spread chars</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
