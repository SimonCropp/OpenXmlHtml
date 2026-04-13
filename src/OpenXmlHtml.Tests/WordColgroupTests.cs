[TestFixture]
public class WordColgroupTests
{
    [Test]
    public Task ColWidth() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col width="100">
              <col width="200">
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task ColWidthPx() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col style="width: 100px">
              <col style="width: 200px">
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task ColWidthInches() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col style="width: 1in">
              <col style="width: 2in">
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task ColSpan() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col span="2" width="100">
              <col width="300">
              <tr><td>A</td><td>B</td><td>C</td></tr>
            </table>
            """));

    [Test]
    public Task Colgroup() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <colgroup>
                <col width="100">
                <col width="200">
              </colgroup>
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task ColgroupWithSpan() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <colgroup span="3" width="150"></colgroup>
              <tr><td>A</td><td>B</td><td>C</td></tr>
            </table>
            """));

    [Test]
    public Task ColgroupMixedWithLooseCol() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <colgroup>
                <col width="100">
              </colgroup>
              <col width="200">
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task ColWidthWithColspan() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col width="100">
              <col width="200">
              <col width="300">
              <tr><td colspan="2">Merged</td><td>C</td></tr>
              <tr><td>A</td><td>B</td><td>C</td></tr>
            </table>
            """));

    [Test]
    public Task CellCssWidthOverridesCol() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col width="100">
              <col width="200">
              <tr><td style="width: 500px">A</td><td>B</td></tr>
            </table>
            """));

    [Test]
    public Task PartialColumnWidths() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col width="100">
              <tr><td>A</td><td>B</td><td>C</td></tr>
            </table>
            """));

    [Test]
    public Task PercentageWidthIgnored() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <table>
              <col width="50%">
              <col width="50%">
              <tr><td>A</td><td>B</td></tr>
            </table>
            """));
}
