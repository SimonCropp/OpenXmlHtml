[TestFixture]
public class WordTextShadowTests
{
    [Test]
    public Task TextShadow() =>
        Verify(WordHtmlConverter.ToParagraphs(
            """<p><span style="text-shadow: 1px 1px 2px black">shadowed text</span></p>"""));

    [Test]
    public Task TextShadowNone() =>
        Verify(WordHtmlConverter.ToParagraphs(
            """<p><span style="text-shadow: none">no shadow</span></p>"""));

    [Test]
    public Task TextShadowConvertToDocx()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-shadow: 2px 2px 4px gray">shadow run</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
