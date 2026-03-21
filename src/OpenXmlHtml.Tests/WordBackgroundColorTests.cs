[TestFixture]
public class WordBackgroundColorTests
{
    [Test]
    public Task RunBackgroundColor()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Normal text with <span style="background-color: #FFFF00">highlighted span</span> inline.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RunBackgroundColorNamed()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Text with <span style="background-color: yellow">yellow</span> and <span style="background-color: lightblue">lightblue</span> backgrounds.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RunBackgroundShorthand()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Using <span style="background: #90EE90">background shorthand</span> property.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ParagraphBackgroundColor()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <p style="background-color: #F0F0F0">Paragraph with gray background</p>
            <p>Normal paragraph</p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ParagraphBackgroundWithFormatting()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <div style="background-color: #FFF3CD; padding: 12pt; margin: 6pt">
              <p><b>Warning:</b> This is a highlighted callout box with padding and margin.</p>
            </div>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MarkElement()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Please review the <mark>important section</mark> before proceeding.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MarkWithOtherFormatting()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>This has <b><mark>bold highlighted</mark></b> and <i><mark>italic highlighted</mark></i> text.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RunAndParagraphBackground()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <p style="background-color: #E8F4FD">
              Light blue paragraph with <span style="background-color: #FFFF00">yellow highlighted</span> text inside.
            </p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
