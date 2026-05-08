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

    [Test]
    public void DirectionRtlPaddedValue()
    {
        var padded = WordHtmlConverter.ToElements("""<p style="direction:   rtl   ">x</p>""");
        var unpadded = WordHtmlConverter.ToElements("""<p style="direction: rtl">x</p>""");
        Assert.That(Xml(padded), Is.EqualTo(Xml(unpadded)));
    }

    [Test]
    public void VerticalRlPaddedValue()
    {
        var padded = WordHtmlConverter.ToElements("""<p style="writing-mode:   vertical-rl   ">x</p>""");
        var unpadded = WordHtmlConverter.ToElements("""<p style="writing-mode: vertical-rl">x</p>""");
        Assert.That(Xml(padded), Is.EqualTo(Xml(unpadded)));
    }

    [Test]
    public void CellWritingModePaddedValue()
    {
        var padded = WordHtmlConverter.ToElements(
            """<table><tr><td style="writing-mode:   vertical-rl   ">x</td></tr></table>""");
        var unpadded = WordHtmlConverter.ToElements(
            """<table><tr><td style="writing-mode: vertical-rl">x</td></tr></table>""");
        Assert.That(Xml(padded), Is.EqualTo(Xml(unpadded)));
    }

    static string Xml(List<OpenXmlElement> elements) =>
        string.Join('\n', elements.Select(e => e.OuterXml));
}
