[TestFixture]
public class SpreadsheetTableTests
{
    [Test]
    public Task SimpleTable() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<table><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></table>"));

    [Test]
    public Task TableWithHeaders() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            """
            <table>
              <thead>
                <tr><th>Name</th><th>Value</th></tr>
              </thead>
              <tbody>
                <tr><td>foo</td><td>bar</td></tr>
              </tbody>
            </table>
            """));

    [Test]
    public Task TableWithCaption() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<table><caption>Table 1</caption><tr><td>data</td></tr></table>"));

    [Test]
    public Task TableWithTfoot() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            """
            <table>
              <tbody>
                <tr><td>row</td></tr>
              </tbody>
              <tfoot>
                <tr><td>total</td></tr>
              </tfoot>
            </table>
            """));

    [Test]
    public Task SingleCellTable() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<table><tr><td>only cell</td></tr></table>"));

    [Test]
    public Task FormattedCellContent() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<table><tr><td><b>bold</b></td><td><i>italic</i></td></tr></table>"));

    [Test]
    public Task ThreeCols() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<table><tr><td>A</td><td>B</td><td>C</td></tr></table>"));

    [Test]
    public Task ColElement() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString(
            "<table><col><col><tr><td>A</td><td>B</td></tr></table>"));
}
