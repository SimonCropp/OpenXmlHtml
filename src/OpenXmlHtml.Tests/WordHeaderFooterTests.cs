[TestFixture]
public class WordHeaderFooterTests
{
    [Test]
    public Task SimpleHeader()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);

        WordHtmlConverter.AppendHtml(body, "<p>Document body content</p>", main);
        WordHtmlConverter.SetHeader(main, "<p>Page Header</p>");

        document.Dispose();
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task SimpleFooter()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);

        WordHtmlConverter.AppendHtml(body, "<p>Document body content</p>", main);
        WordHtmlConverter.SetFooter(main, "<p>Page Footer</p>");

        document.Dispose();
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task HeaderAndFooter()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);

        WordHtmlConverter.AppendHtml(body, "<h1>Report</h1><p>Content here.</p>", main);
        WordHtmlConverter.SetHeader(main,
            """<p style="text-align: center"><b>Company Name</b> — Confidential</p>""");
        WordHtmlConverter.SetFooter(main,
            """<p style="text-align: center; font-size: 9pt; color: gray">Page footer — Generated 2024</p>""");

        document.Dispose();
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task FormattedHeader()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);

        WordHtmlConverter.AppendHtml(body, "<p>Body</p>", main);
        WordHtmlConverter.SetHeader(main,
            """
            <table style="width: 500pt; border: none">
              <tr>
                <td style="border: none"><b>ACME Corp</b></td>
                <td style="border: none; text-align: right"><i>Internal Use Only</i></td>
              </tr>
            </table>
            <hr>
            """);

        document.Dispose();
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task FormattedFooter()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);

        WordHtmlConverter.AppendHtml(body, "<p>Body</p>", main);
        WordHtmlConverter.SetFooter(main,
            """
            <hr>
            <p style="font-size: 8pt; color: #666666; text-align: center">
              © 2024 ACME Corp. All rights reserved. | <a href="https://example.com">Privacy Policy</a>
            </p>
            """);

        // Verify body + footer XML (not full docx binary) to avoid
        // non-deterministic relationship IDs in footer .rels
        // (DeterministicIoPackaging 0.24.3+ will fix this)
        var footer = main.FooterParts.First().Footer!;
        return Verify(
            new
            {
                body,
                footer
            });
    }
}
