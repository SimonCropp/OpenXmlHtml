[TestFixture]
public class WordConvertToDocxTests
{
    [Test]
    public Task SimpleHtml()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx("<p>Hello <b>World</b></p>", stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task RichDocument()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<h1>Report</h1>" +
            "<p>Generated on <time>2024-01-15</time></p>" +
            "<h2>Summary</h2>" +
            "<p>All <span style=\"color: green\"><b>systems operational</b></span>.</p>" +
            "<ul><li>Server: OK</li><li>Database: OK</li></ul>" +
            "<h2>Details</h2>" +
            "<p>See <a href=\"https://example.com\">report</a> for more.</p>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task FormattedContent()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<p><b>Bold</b>, <i>italic</i>, <u>underline</u>, <s>strike</s></p>" +
            "<p><sup>super</sup> and <sub>sub</sub></p>" +
            "<p><code>monospace</code> and <small>small</small></p>" +
            "<p><font color=\"#FF0000\" size=\"18\" face=\"Arial\">styled font</font></p>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task TableContent()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<table>" +
            "<caption>Q1 Sales</caption>" +
            "<thead><tr><th>Region</th><th>Revenue</th></tr></thead>" +
            "<tbody>" +
            "<tr><td>North</td><td>$500K</td></tr>" +
            "<tr><td>South</td><td>$700K</td></tr>" +
            "</tbody>" +
            "</table>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task OrderedAndUnorderedLists()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<h3>Unordered</h3>" +
            "<ul><li>Alpha</li><li>Beta</li><li>Gamma</li></ul>" +
            "<h3>Ordered</h3>" +
            "<ol><li>First</li><li>Second</li><li>Third</li></ol>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task PreformattedAndBlockquote()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<blockquote>To be or not to be</blockquote>" +
            "<pre>  line 1\n  line 2\n  line 3</pre>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task EmptyHtml()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx("", stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task LinksAndImages()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<p>Visit <a href=\"https://example.com\">our site</a></p>" +
            "<figure>" +
            "<img alt=\"Company Logo\">" +
            "<figcaption>Figure 1: Logo</figcaption>" +
            "</figure>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task CssStyles()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<p style=\"font-weight: bold; color: navy\">Navy bold</p>" +
            "<p><span style=\"font-family: Verdana; font-size: 14pt\">Verdana 14pt</span></p>" +
            "<p><span style=\"text-decoration: underline line-through\">Both decorations</span></p>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task DefinitionList()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<dl>" +
            "<dt>HTML</dt><dd>HyperText Markup Language</dd>" +
            "<dt>CSS</dt><dd>Cascading Style Sheets</dd>" +
            "</dl>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task InlineQuotation()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<p>She said <q>hello world</q> to everyone.</p>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    [Test]
    public Task SemanticElements()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            "<article>" +
            "<header><h1>Article Title</h1></header>" +
            "<section><p>Main content with <abbr title=\"HyperText Markup Language\">HTML</abbr>.</p></section>" +
            "<aside><p>Related info</p></aside>" +
            "<footer><p><small>Copyright 2024</small></p></footer>" +
            "</article>",
            stream);
        stream.Position = 0;
        return VerifyStream(stream, "docx");
    }

    static SettingsTask VerifyStream(Stream stream, string extension) =>
        Verify(stream, extension);
}
