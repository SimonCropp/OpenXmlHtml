[TestFixture]
public class WordTableTests
{
    [Test]
    public Task SimpleTable() =>
        Verify(WordHtmlConverter.ToElements(
            "<table><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></table>"));

    [Test]
    public Task TableWithHeaders() =>
        Verify(WordHtmlConverter.ToElements(
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
        Verify(WordHtmlConverter.ToElements(
            "<table><caption>Table 1</caption><tr><td>data</td></tr></table>"));

    [Test]
    public Task FormattedCellContent() =>
        Verify(WordHtmlConverter.ToElements(
            "<table><tr><td><b>bold</b></td><td><i>italic</i></td></tr></table>"));

    [Test]
    public Task TableWithColspan() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <tr><td colspan="2">Merged</td></tr>
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task TableWithRowspan() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <tr><td rowspan="2">Span</td><td>B1</td></tr>
              <tr><td>B2</td></tr>
            </table>
            """));

    [Test]
    public Task NestedTable() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <tr>
                <td>Outer</td>
                <td>
                  <table><tr><td>Inner</td></tr></table>
                </td>
              </tr>
            </table>
            """));

    [Test]
    public Task MixedContentWithTable() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <p>Before table</p>
            <table><tr><td>Cell</td></tr></table>
            <p>After table</p>
            """));

    [Test]
    public Task TableWithTfoot() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <thead><tr><th>Header</th></tr></thead>
              <tbody><tr><td>Body</td></tr></tbody>
              <tfoot><tr><td>Footer</td></tr></tfoot>
            </table>
            """));

    [Test]
    public Task EmptyTable() =>
        Verify(WordHtmlConverter.ToElements("<table></table>"));
}
