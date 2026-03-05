[TestFixture]
public class WordSamples
{
    [Test]
    public Task ToParagraphs()
    {
        #region ToParagraphs

        var paragraphs = WordHtmlConverter.ToParagraphs(
            """
            <h1>Report Title</h1>
            <p>This is a <b>bold</b> statement with <i>emphasis</i>.</p>
            """);

        #endregion

        return Verify(paragraphs);
    }

    [Test]
    public Task AppendHtml()
    {
        #region AppendHtml

        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(
            stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new(new Body());

        WordHtmlConverter.AppendHtml(
            mainPart.Document.Body!,
            """
            <h1>Meeting Notes</h1>
            <p><i>Date: January 15, 2024</i></p>
            <ol>
              <li>Review <code>PR #123</code></li>
              <li>Update <u>documentation</u></li>
            </ol>
            """);

        #endregion

        return Verify(mainPart.Document.Body!);
    }

    [Test]
    public Task RichDocument()
    {
        #region WordRichDocument

        var paragraphs = WordHtmlConverter.ToParagraphs(
            """
            <h2>Status Report</h2>
            <p>All systems <span style="color: green"><b>operational</b></span>.</p>
            <ul>
              <li>Server: <span style="color: green">OK</span></li>
              <li>Cache: <span style="color: red">Down</span></li>
            </ul>
            <p>Contact <a href="mailto:ops@example.com">ops team</a> for details.</p>
            """);

        #endregion

        return Verify(paragraphs);
    }

    [Test]
    public Task ConvertToDocx()
    {
        #region ConvertToDocx

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <h1>Report</h1>
            <p>This is a <b>bold</b> statement.</p>
            <ul>
              <li>Item one</li>
              <li>Item two</li>
            </ul>
            """,
            stream);

        #endregion

        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ConvertStreamToDocx()
    {
        #region ConvertStreamToDocx

        using var htmlStream = new MemoryStream(
            "<h1>Report</h1><p>Content</p>"u8.ToArray());
        using var docxStream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(htmlStream, docxStream);

        #endregion

        docxStream.Position = 0;
        return Verify(docxStream, "docx");
    }

    [Test]
    public async Task ConvertFileToDocx()
    {
        var htmlPath = await TempFile.CreateText("<h1>Hello</h1><p>World</p>");
        var docxPath = new TempFile("docx");

        #region ConvertFileToDocx

        WordHtmlConverter.ConvertFileToDocx(htmlPath, docxPath);

        #endregion

        await VerifyFile(docxPath);
    }
}
