[TestFixture]
public class WordBorderTests
{
    [Test]
    public Task RunBorderSolid()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Text with <span style="border: 1px solid black">bordered span</span> inline.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RunBorderDotted()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Text with <span style="border: 1px dotted red">dotted border</span> inline.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RunBorderDashed()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Text with <span style="border: 2px dashed #0000FF">dashed blue border</span> inline.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RunBorderDouble()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p>Text with <span style="border: 2px double green">double border</span> inline.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ParagraphBorderAllSides()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="border: 1px solid #333333; padding: 6pt">Bordered paragraph</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ParagraphBorderIndividualSides()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="border-left: 3px solid blue; border-bottom: 1px dashed gray">Left and bottom borders</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ParagraphBorderBottomOnly()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="border-bottom: 2px solid black">Underlined paragraph</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CellBorder()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="border: 2px solid red">Red bordered cell</td>
                <td>Normal cell</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CellBorderIndividualSides()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="border-bottom: 3px double blue">Bottom double border</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TableBorderOverride()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table style="border: 2px dashed red">
              <tr>
                <td>Dashed red table border</td>
                <td>Inherited by all cells</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TableBorderNone()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table style="border: none">
              <tr>
                <td>No borders</td>
                <td>Borderless table</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HtmlBorderAttributeZero()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table border="0">
              <tr>
                <td>No borders via HTML attr</td>
                <td>border="0"</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HtmlBorderAttributeOne()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table border="1">
              <tr>
                <td>Cell A</td>
                <td>Cell B</td>
              </tr>
              <tr>
                <td>Cell C</td>
                <td>Cell D</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HtmlBorderAttributeThick()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table border="3">
              <tr>
                <td>Thick</td>
                <td>Borders</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task BorderWithBackground()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <div style="border: 1px solid #CCC; background-color: #FFFDE7; padding: 12pt; margin: 6pt">
              <p><b>Note:</b> This is a callout box with border and background.</p>
            </div>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task BorderStyleVariants()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <p style="border: 1px groove gray">Groove border</p>
            <p style="border: 1px ridge gray">Ridge border</p>
            <p style="border: 1px inset gray">Inset border</p>
            <p style="border: 1px outset gray">Outset border</p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
