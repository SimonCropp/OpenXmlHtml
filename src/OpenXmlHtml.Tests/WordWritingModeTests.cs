[TestFixture]
public class WordWritingModeTests
{
    [Test]
    public Task VerticalRl()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="writing-mode: vertical-rl">Vertical right-to-left text</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task VerticalLr()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="writing-mode: vertical-lr">Vertical left-to-right text</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task DirectionRtl()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="direction: rtl">Right-to-left direction text</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task VerticalInTableCell()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="writing-mode: vertical-rl; height: 100pt">Vertical cell text</td>
                <td>Normal cell text</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HorizontalDefault()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="writing-mode: horizontal-tb">Default horizontal text</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
