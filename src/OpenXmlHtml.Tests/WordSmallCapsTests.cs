[TestFixture]
public class WordSmallCapsTests
{
    [Test]
    public Task SmallCapsCss()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Normal text with <span style="font-variant: small-caps">small caps text</span> inline.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task SmallCapsWithOtherFormatting()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><b style="font-variant: small-caps">Bold Small Caps</b> and <i style="font-variant: small-caps">Italic Small Caps</i></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
