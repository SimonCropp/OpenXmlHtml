using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

#pragma warning disable CA1822 // Mark members as static - BenchmarkDotNet requires instance methods

namespace OpenXmlHtml.Benchmarks;

[MemoryDiagnoser]
public class WordBenchmarks
{
    const string simpleInline = "<b>Bold</b> and <i>italic</i> with <span style=\"color: red\">color</span>";

    const string richParagraphs = """
        <h1>Report</h1>
        <p>This is a <b>bold</b> and <i>italic</i> paragraph with <a href="https://example.com">a link</a>.</p>
        <p style="margin-left: 36pt; text-indent: 18pt; line-height: 1.5">
          Indented paragraph with <span style="color: navy; font-size: 14pt; font-family: Georgia">styled span</span>
          and <code>inline code</code> and <mark>highlighted text</mark>.
        </p>
        <blockquote>A blockquote with <b>bold</b> content.</blockquote>
        <pre>Preformatted
          text block</pre>
        """;

    const string nestedList = """
        <ol>
          <li><b>First</b> item
            <ul>
              <li>Nested bullet <i>italic</i></li>
              <li>Another bullet with <a href="https://example.com">link</a>
                <ol>
                  <li>Deep nested ordered</li>
                  <li>Another deep item</li>
                </ol>
              </li>
            </ul>
          </li>
          <li>Second item</li>
          <li><del>Removed</del> <ins>Added</ins></li>
        </ol>
        """;

    const string table = """
        <table style="border: 1px solid black; width: 100%">
          <caption>Quarterly Results</caption>
          <thead>
            <tr><th>Quarter</th><th>Revenue</th><th>Growth</th></tr>
          </thead>
          <tbody>
            <tr>
              <td>Q1</td>
              <td style="text-align: right; background-color: #eee">$12.3M</td>
              <td><span style="color: green; font-weight: bold">+8%</span></td>
            </tr>
            <tr>
              <td>Q2</td>
              <td style="text-align: right; background-color: #eee">$14.1M</td>
              <td><span style="color: green; font-weight: bold">+12%</span></td>
            </tr>
            <tr>
              <td colspan="2" style="font-weight: bold">Total</td>
              <td><b>+10%</b></td>
            </tr>
          </tbody>
        </table>
        """;

    const string nestedTable = """
        <table>
          <tr>
            <td>Outer</td>
            <td>
              <table>
                <tr><td>Inner A</td><td>Inner B</td></tr>
                <tr><td>Inner C</td><td>Inner D</td></tr>
              </table>
            </td>
          </tr>
        </table>
        """;

    static readonly string largeDocument;

    static WordBenchmarks()
    {
        // Build a large document by repeating sections
        var sections = new System.Text.StringBuilder();
        for (var i = 0; i < 20; i++)
        {
            sections.Append($"""
                <h2>Section {i + 1}</h2>
                {richParagraphs}
                {nestedList}
                {table}
                """);
        }

        largeDocument = sections.ToString();
    }

    // --- ToParagraphs (flat segment path) ---

    [Benchmark]
    public List<Paragraph> ToParagraphs_SimpleInline() =>
        WordHtmlConverter.ToParagraphs(simpleInline);

    [Benchmark]
    public List<Paragraph> ToParagraphs_RichParagraphs() =>
        WordHtmlConverter.ToParagraphs(richParagraphs);

    // --- ToElements (DOM-based path, no MainDocumentPart) ---

    [Benchmark]
    public List<OpenXmlElement> ToElements_SimpleInline() =>
        WordHtmlConverter.ToElements(simpleInline);

    [Benchmark]
    public List<OpenXmlElement> ToElements_RichParagraphs() =>
        WordHtmlConverter.ToElements(richParagraphs);

    [Benchmark]
    public List<OpenXmlElement> ToElements_NestedList() =>
        WordHtmlConverter.ToElements(nestedList);

    [Benchmark]
    public List<OpenXmlElement> ToElements_Table() =>
        WordHtmlConverter.ToElements(table);

    [Benchmark]
    public List<OpenXmlElement> ToElements_NestedTable() =>
        WordHtmlConverter.ToElements(nestedTable);

    // --- ConvertToDocx (full pipeline with MainDocumentPart) ---

    [Benchmark]
    public void ConvertToDocx_RichParagraphs()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(richParagraphs, stream);
    }

    [Benchmark]
    public void ConvertToDocx_Table()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(table, stream);
    }

    [Benchmark]
    public void ConvertToDocx_NestedList()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(nestedList, stream);
    }

    [Benchmark]
    public void ConvertToDocx_LargeDocument()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(largeDocument, stream);
    }

    // --- AppendHtml (into existing document) ---

    [Benchmark]
    public void AppendHtml_RichParagraphs()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();
        mainPart.Document = new(body);
        WordHtmlConverter.AppendHtml(body, richParagraphs, mainPart);
    }

    [Benchmark]
    public void AppendHtml_LargeDocument()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();
        mainPart.Document = new(body);
        WordHtmlConverter.AppendHtml(body, largeDocument, mainPart);
    }
}
