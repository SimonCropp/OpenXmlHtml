[TestFixture]
public class WordWhiteSpaceTests
{
    [Test]
    public Task WhiteSpacePre() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<div style=\"white-space: pre\">  spaces   preserved\n  and  newlines</div>"));

    [Test]
    public Task WhiteSpacePreWrap() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<div style=\"white-space: pre-wrap\">  multiple   spaces  </div>"));

    [Test]
    public Task WhiteSpaceNowrap() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<p style=\"white-space: nowrap\">no breaks here please</p>"));

    [Test]
    public Task WhiteSpaceNormal() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<div style=\"white-space: normal\">   collapsed   spaces   </div>"));

    [Test]
    public Task WhiteSpaceConvertToDocx()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<p style=\"white-space: pre\">indented  text  preserved</p>",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
