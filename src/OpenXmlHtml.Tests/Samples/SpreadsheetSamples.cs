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

    [Test]
    public Task BugReportCell()
    {
        var cell = new SpreadsheetCell();
        SpreadsheetHtmlConverter.SetCellHtml(cell,
            """
            <h3>BUG-4821: Login fails after password reset</h3>
            <p><b>Priority:</b> <span style="color: red">Critical</span>
            &nbsp;|&nbsp; <b>Assignee:</b> <i>J. Martinez</i></p>

            <p><b>Steps to reproduce:</b></p>
            <ol>
              <li>Reset password via <a href="https://auth.example.com/reset">auth portal</a></li>
              <li>Attempt login with new credentials</li>
              <li>Observe <span style="color: red"><b>403 Forbidden</b></span> error</li>
            </ol>

            <p><b>Expected:</b> <span style="color: green">Successful login</span><br>
            <b>Actual:</b> <s>Session token refreshed</s> → <code>ERR_AUTH_STALE_TOKEN</code></p>

            <p><b>Environment:</b></p>
            <ul>
              <li>Browser: Chrome 120<sup>*</sup></li>
              <li>OS: Windows 11</li>
              <li>API version: <font face="Courier New" size="10">v2.4.1</font></li>
            </ul>

            <p><small>* Also reproduced on Firefox and Safari.</small></p>

            <p><b>Notes:</b> Likely related to the token cache changes in
            <a href="https://jira.example.com/PR-891">PR-891</a>.
            The <kbd>session_id</kbd> cookie is set but the
            H<sub>2</sub>O auth header contains a stale nonce.</p>
            """);
        return Verify(cell);
    }

    [Test]
    public Task BugReportSpreadsheet()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new(sheetData);

            workbookPart.Workbook = new(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Bug Report"
                    }));

            // Header row
            var headerRow = new Row
            {
                RowIndex = 1
            };
            string[] headers = ["ID", "Priority", "Summary", "Details"];
            foreach (var header in headers)
            {
                var cell = new SpreadsheetCell
                {
                    DataType = CellValues.InlineString,
                    InlineString = SpreadsheetHtmlConverter.ToInlineString($"<b>{header}</b>")
                };
                headerRow.Append(cell);
            }

            sheetData.Append(headerRow);

            // Bug row
            var bugRow = new Row
            {
                RowIndex = 2
            };

            var idCell = new SpreadsheetCell();
            SpreadsheetHtmlConverter.SetCellHtml(idCell, "<code>BUG-4821</code>");
            bugRow.Append(idCell);

            var priorityCell = new SpreadsheetCell();
            SpreadsheetHtmlConverter.SetCellHtml(priorityCell,
                """<span style="color: red"><b>Critical</b></span>""");
            bugRow.Append(priorityCell);

            var summaryCell = new SpreadsheetCell();
            SpreadsheetHtmlConverter.SetCellHtml(summaryCell,
                "Login fails after <u>password reset</u> — returns <s>200</s> <b>403</b>");
            bugRow.Append(summaryCell);

            var detailsCell = new SpreadsheetCell();
            SpreadsheetHtmlConverter.SetCellHtml(detailsCell,
                """
                <p><b>Steps:</b></p>
                <ol>
                  <li>Reset password via <a href="https://auth.example.com/reset">auth portal</a></li>
                  <li>Attempt login with new credentials</li>
                  <li>Observe <span style="color: red"><b>403 Forbidden</b></span></li>
                </ol>
                <p><b>Expected:</b> <span style="color: green">Successful login</span><br>
                <b>Actual:</b> <code>ERR_AUTH_STALE_TOKEN</code></p>
                <p><small>Affects Chrome 120<sup>*</sup>, Firefox, Safari.
                Related to <a href="https://jira.example.com/PR-891">PR-891</a>.
                The H<sub>2</sub>O auth header contains a stale nonce.</small></p>
                """);
            bugRow.Append(detailsCell);

            sheetData.Append(bugRow);
        }

        stream.Position = 0;
        return Verify(stream, "xlsx");
    }
}
