using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

#pragma warning disable CA1822 // Mark members as static - BenchmarkDotNet requires instance methods

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
    static readonly string hyperlinkHeavy;
    static readonly string lineHeightHeavy;
    static readonly string directionRtlHeavy;
    static readonly string svgDimensionHeavy;

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

        // Targeted hot-path inputs: dense repetition of a single feature so
        // parser-level fixed costs are dwarfed by the per-element allocations
        // touched by the span-conversion changes.

        // Inputs use padded whitespace where the parser does not pre-trim the value,
        // so the trim allocation difference between string.Trim() (allocates) and
        // span.Trim() (no allocation) shows up.

        var links = new System.Text.StringBuilder();
        for (var i = 0; i < 100; i++)
        {
            // TextContent preserves leading/trailing whitespace inside <a>.
            links.Append($"<p><a href=\"https://example.com/page{i}\">  link {i}  </a></p>");
        }
        hyperlinkHeavy = links.ToString();

        // StyleParser.Parse pre-trims values from inline style="...", so the
        // ParagraphFormatState span changes for line-height and direction do
        // not have any input that would make Trim() allocate. Kept as a
        // regression-witness that the changes did not slow the path down.
        var lineHeights = new System.Text.StringBuilder();
        for (var i = 0; i < 100; i++)
        {
            lineHeights.Append($"<p style=\"line-height: 1.5\">paragraph {i}</p>");
        }
        lineHeightHeavy = lineHeights.ToString();

        var rtl = new System.Text.StringBuilder();
        for (var i = 0; i < 100; i++)
        {
            rtl.Append($"<p style=\"direction: rtl\">paragraph {i}</p>");
        }
        directionRtlHeavy = rtl.ToString();

        // Raw HTML width/height attributes are NOT pre-trimmed, so padded
        // values exercise both the bug fix (EndsWith on trimmed span) and the
        // Substring-replaced-with-span allocation savings.
        var svg = new System.Text.StringBuilder();
        for (var i = 0; i < 100; i++)
        {
            svg.Append("<svg width=\"100px \" height=\" 50px\" viewBox=\"0 0 100 50\"><rect width=\"10\" height=\"10\"/></svg>");
        }
        svgDimensionHeavy = svg.ToString();
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
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);
        WordHtmlConverter.AppendHtml(body, richParagraphs, main);
    }

    [Benchmark]
    public void AppendHtml_LargeDocument()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        main.Document = new(body);
        WordHtmlConverter.AppendHtml(body, largeDocument, main);
    }

    // --- Targeted hot-path benchmarks (span-conversion changes) ---

    // HtmlSegmentParser.ProcessNode "a" branch — hits HtmlParser.cs:144 linkText comparison.
    [Benchmark]
    public List<Paragraph> ToParagraphs_HyperlinkHeavy() =>
        WordHtmlConverter.ToParagraphs(hyperlinkHeavy);

    // WordContentBuilder hyperlink branch — hits WordContentBuilder.cs:566 linkText comparison.
    [Benchmark]
    public List<OpenXmlElement> ToElements_HyperlinkHeavy() =>
        WordHtmlConverter.ToElements(hyperlinkHeavy);

    // ParagraphFormatState.ParseLineHeight — hits the lhSpan TryParse/Contains path.
    [Benchmark]
    public List<OpenXmlElement> ToElements_LineHeightHeavy() =>
        WordHtmlConverter.ToElements(lineHeightHeavy);

    // ParagraphFormatState.From — hits the direction.AsSpan().Trim().Equals path.
    [Benchmark]
    public List<OpenXmlElement> ToElements_DirectionRtlHeavy() =>
        WordHtmlConverter.ToElements(directionRtlHeavy);

    // HtmlParser.ParseSvgDimension — hits the trimmed-span EndsWith("px") path.
    [Benchmark]
    public List<OpenXmlElement> ToElements_SvgDimensionHeavy() =>
        WordHtmlConverter.ToElements(svgDimensionHeavy);
}
