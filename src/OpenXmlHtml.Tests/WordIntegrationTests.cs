[TestFixture]
public class WordIntegrationTests
{
    [Test]
    public Task AppendHtmlToBody()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        WordHtmlConverter.AppendHtml(mainPart.Document.Body!,
            "<h1>Report Title</h1><p>This is a <b>bold</b> statement.</p>");

        return Verify(mainPart.Document.Body!);
    }

    [Test]
    public Task RichFormattedDocument()
    {
        var paragraphs = WordHtmlConverter.ToParagraphs(
            "<h1>Status Report</h1>" +
            "<p><b>Date:</b> 2024-01-15</p>" +
            "<p>All systems <span style=\"color: green\"><b>operational</b></span>.</p>" +
            "<ul><li>Server: <span style=\"color: green\">OK</span></li>" +
            "<li>Database: <span style=\"color: green\">OK</span></li>" +
            "<li>Cache: <span style=\"color: red\">Down</span></li></ul>");

        return Verify(paragraphs);
    }

    [Test]
    public Task FormattedReport()
    {
        var paragraphs = WordHtmlConverter.ToParagraphs(
            "<h2>Meeting Notes</h2>" +
            "<p><i>Date: January 15, 2024</i></p>" +
            "<p>Attendees: <b>Alice</b>, <b>Bob</b>, <b>Charlie</b></p>" +
            "<h3>Action Items</h3>" +
            "<ol><li>Review <code>PR #123</code></li>" +
            "<li>Update <u>documentation</u></li>" +
            "<li><del>Fix bug #456</del> <ins>Done!</ins></li></ol>");

        return Verify(paragraphs);
    }

    [Test]
    public Task MultiParagraphWithBreaks()
    {
        var paragraphs = WordHtmlConverter.ToParagraphs(
            "Line 1<br>Line 2<br>Line 3");

        return Verify(paragraphs);
    }
}
