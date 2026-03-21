[TestFixture]
public class WordUnderlineTests
{
    [Test]
    public Task SingleUnderline()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration: underline">Single underline</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task DottedUnderline()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration: underline; text-decoration-style: dotted">Dotted underline</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task DashedUnderline()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration: underline; text-decoration-style: dashed">Dashed underline</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task WavyUnderline()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration: underline; text-decoration-style: wavy">Wavy underline</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task DoubleUnderline()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration: underline; text-decoration-style: double">Double underline</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task UnderlineFromTag()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<p><u>Tag underline</u> and <ins>ins underline</ins></p>",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task UnderlineWithColor()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration: underline; text-decoration-style: wavy; color: red">Wavy red underline</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task DecorationStyleWithoutUnderline_NoEffect()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><span style="text-decoration-style: dotted">No underline set, style should be ignored</span></p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task UnderlineOnSpreadsheet() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            """<span style="text-decoration: underline; text-decoration-style: dotted">Underlined in spreadsheet</span>"""));
}
