[TestFixture]
public class WordTableStyleTests
{
    [Test]
    public Task CellPadding()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="padding: 12pt">Padded cell</td>
                <td>Normal cell</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CellPaddingIndividual()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="padding-top: 6pt; padding-bottom: 12pt; padding-left: 24pt; padding-right: 8pt">Custom padding</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CellWidth()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="width: 200pt">Wide cell</td>
                <td style="width: 100pt">Narrow cell</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CellBackgroundColor()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <th style="background-color: #4472C4; color: white">Header</th>
                <th style="background-color: #4472C4; color: white">Value</th>
              </tr>
              <tr>
                <td>Row 1</td>
                <td style="background-color: #E2EFDA">$100</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CellVerticalAlign()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td style="vertical-align: top; padding: 12pt">Top</td>
                <td style="vertical-align: middle; padding: 12pt">Middle</td>
                <td style="vertical-align: bottom; padding: 12pt">Bottom</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TableWidth()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table style="width: 400pt">
              <tr>
                <td>Cell 1</td>
                <td>Cell 2</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TableBackgroundColor()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table style="background-color: #F2F2F2">
              <tr>
                <td>Shaded table</td>
                <td>Background from table</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TableDefaultCellPadding()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table style="padding: 8pt">
              <tr>
                <td>Default padding from table</td>
                <td>Applied to all cells</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HtmlCellpaddingAttribute()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table cellpadding="10">
              <tr>
                <td>Cell with HTML cellpadding</td>
                <td>Applied via attribute</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HtmlBgcolorAttribute()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td bgcolor="#FFCCCC">Pink via bgcolor attr</td>
                <td>Normal cell</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HtmlWidthAttribute()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr>
                <td width="200">Wide via attr</td>
                <td width="100">Narrow via attr</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CombinedCellStyles()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table style="width: 500pt">
              <tr>
                <th style="background-color: #2E4057; color: white; padding: 8pt; width: 150pt">Product</th>
                <th style="background-color: #2E4057; color: white; padding: 8pt; width: 100pt">Price</th>
                <th style="background-color: #2E4057; color: white; padding: 8pt">Status</th>
              </tr>
              <tr>
                <td style="padding: 6pt">Widget A</td>
                <td style="padding: 6pt; text-align: right">$29.99</td>
                <td style="padding: 6pt; background-color: #E2EFDA; vertical-align: middle">
                  <span style="color: green"><b>In Stock</b></span>
                </td>
              </tr>
              <tr>
                <td style="padding: 6pt">Widget B</td>
                <td style="padding: 6pt; text-align: right">$49.99</td>
                <td style="padding: 6pt; background-color: #FCE4EC; vertical-align: middle">
                  <span style="color: red"><b>Out of Stock</b></span>
                </td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public void CellVerticalAlignPaddedValue()
    {
        var padded = WordHtmlConverter.ToElements(
            """<table><tr><td style="vertical-align:   top   ">x</td></tr></table>""");
        var unpadded = WordHtmlConverter.ToElements(
            """<table><tr><td style="vertical-align: top">x</td></tr></table>""");
        Assert.That(
            string.Join('\n', padded.Select(e => e.OuterXml)),
            Is.EqualTo(string.Join('\n', unpadded.Select(e => e.OuterXml))));
    }
}
