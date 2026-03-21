[TestFixture]
public class WordRowHeightTests
{
    [Test]
    public Task RowHeightCss()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr style="height: 50pt">
                <td>Tall row via CSS</td>
              </tr>
              <tr>
                <td>Normal row</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RowHeightAttribute()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <table>
              <tr height="80">
                <td>Tall row via attribute</td>
              </tr>
              <tr>
                <td>Normal row</td>
              </tr>
            </table>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
