[TestFixture]
public class WordSamples
{
    [Test]
    public void ToParagraphs()
    {
        #region ToParagraphs

        var paragraphs = WordHtmlConverter.ToParagraphs(
            "<h1>Report Title</h1>" +
            "<p>This is a <b>bold</b> statement with <i>emphasis</i>.</p>");

        #endregion
    }

    [Test]
    public void AppendHtml()
    {
        #region AppendHtml

        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(
            stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        WordHtmlConverter.AppendHtml(
            mainPart.Document.Body!,
            "<h1>Meeting Notes</h1>" +
            "<p><i>Date: January 15, 2024</i></p>" +
            "<ol>" +
            "<li>Review <code>PR #123</code></li>" +
            "<li>Update <u>documentation</u></li>" +
            "</ol>");

        #endregion
    }

    [Test]
    public void RichDocument()
    {
        #region WordRichDocument

        var paragraphs = WordHtmlConverter.ToParagraphs(
            "<h2>Status Report</h2>" +
            "<p>All systems <span style=\"color: green\"><b>operational</b></span>.</p>" +
            "<ul>" +
            "<li>Server: <span style=\"color: green\">OK</span></li>" +
            "<li>Cache: <span style=\"color: red\">Down</span></li>" +
            "</ul>" +
            "<p>Contact <a href=\"mailto:ops@example.com\">ops team</a> for details.</p>");

        #endregion
    }

    [Test]
    public void ConvertToDocx()
    {
        #region ConvertToDocx

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<h1>Report</h1>" +
            "<p>This is a <b>bold</b> statement.</p>" +
            "<ul><li>Item one</li><li>Item two</li></ul>",
            stream);

        #endregion
    }

    [Test]
    public void ConvertFileToDocx()
    {
        var htmlPath = Path.GetTempFileName();
        var docxPath = Path.ChangeExtension(htmlPath, ".docx");
        try
        {
            File.WriteAllText(htmlPath, "<h1>Hello</h1><p>World</p>");

            #region ConvertFileToDocx

            WordHtmlConverter.ConvertFileToDocx(htmlPath, docxPath);

            #endregion
        }
        finally
        {
            File.Delete(htmlPath);
            File.Delete(docxPath);
        }
    }
}
