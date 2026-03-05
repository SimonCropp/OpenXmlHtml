[TestFixture]
public class SpreadsheetIntegrationTests
{
    [Test]
    public Task SetCellHtml()
    {
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        var row = new Row { RowIndex = 1 };
        sheetData.Append(row);

        var cell = new SpreadsheetCell { CellReference = "A1" };
        row.Append(cell);

        SpreadsheetHtmlConverter.SetCellHtml(cell, "<b>Hello</b> <i>World</i>");

        return Verify(cell);
    }

    [Test]
    public Task RichFormattedCell()
    {
        var cell = new SpreadsheetCell();
        SpreadsheetHtmlConverter.SetCellHtml(cell,
            """
            <p>Report <b>Summary</b></p>
            <ul>
              <li><span style="color: red">Critical</span>: 3</li>
              <li><span style="color: green">Passed</span>: 47</li>
            </ul>
            """);
        return Verify(cell);
    }

    [Test]
    public Task ComplexTable()
    {
        var cell = new SpreadsheetCell();
        SpreadsheetHtmlConverter.SetCellHtml(cell,
            """
            <b>Q1 Results</b><br>
            Revenue: <font color="#008000">$1.2M</font><br>
            Expenses: <font color="#FF0000">$800K</font><br>
            <i>Net: <b>$400K</b></i>
            """);
        return Verify(cell);
    }
}
