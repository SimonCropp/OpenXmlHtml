[TestFixture]
public class WordTableTests
{
    [Test]
    public Task SimpleTable() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<table><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></table>"));

    [Test]
    public Task TableWithHeaders() =>
        Verify(WordHtmlConverter.ToParagraphs(
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
        Verify(WordHtmlConverter.ToParagraphs(
            "<table><caption>Table 1</caption><tr><td>data</td></tr></table>"));

    [Test]
    public Task FormattedCellContent() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<table><tr><td><b>bold</b></td><td><i>italic</i></td></tr></table>"));
}
