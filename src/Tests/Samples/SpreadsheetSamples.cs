[TestFixture]
public class SpreadsheetSamples
{
    [Test]
    public Task SetCellHtml()
    {
        #region SetCellHtml

        var cell = new SpreadsheetCell();
        SpreadsheetHtmlConverter.SetCellHtml(cell, "<b>Hello</b> <i>World</i>");

        #endregion

        return Verify(cell);
    }

    [Test]
    public Task ToInlineString()
    {
        #region ToInlineString

        var inlineString = SpreadsheetHtmlConverter.ToInlineString(
            "<b>Revenue:</b> <font color=\"#008000\">$1.2M</font>");

        #endregion

        return Verify(inlineString);
    }

    [Test]
    public Task FormattedList()
    {
        #region SpreadsheetList

        var inlineString = SpreadsheetHtmlConverter.ToInlineString(
            """
            <ul>
              <li><span style="color: green">Passed</span>: 47</li>
              <li><span style="color: red">Failed</span>: 3</li>
            </ul>
            """);

        #endregion

        return Verify(inlineString);
    }

    [Test]
    public Task RichContent()
    {
        #region SpreadsheetRichContent

        var cell = new SpreadsheetCell();
        SpreadsheetHtmlConverter.SetCellHtml(cell,
            """
            <h2>Q1 Report</h2>
            <p>Revenue: <b style="color: green">$1.2M</b></p>
            <p>See <a href="https://example.com/report">full report</a></p>
            <table>
              <tr><th>Region</th><th>Sales</th></tr>
              <tr><td>North</td><td>$500K</td></tr>
              <tr><td>South</td><td>$700K</td></tr>
            </table>
            """);

        #endregion

        return Verify(cell);
    }
}
