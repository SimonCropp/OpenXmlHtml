[TestFixture]
public class WordConvertFileTests
{
    [Test]
    public async Task ConvertHtmlFile()
    {
        var dir = Path.Combine(Path.GetTempPath(), "OpenXmlHtmlTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(dir);
        try
        {
            var htmlPath = Path.Combine(dir, "test.html");
            var docxPath = Path.Combine(dir, "test.docx");

            File.WriteAllText(htmlPath,
                """
                <html>
                <body>
                  <h1>Document Title</h1>
                  <p>A paragraph with <b>bold</b> and <i>italic</i> text.</p>
                  <ul>
                    <li>Item one</li>
                    <li>Item two</li>
                  </ul>
                </body>
                </html>
                """);

            WordHtmlConverter.ConvertFileToDocx(htmlPath, docxPath);

            await VerifyFile(docxPath).UniqueForTargetFrameworkAndVersion();
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Test]
    public async Task ConvertFullHtmlPage()
    {
        var dir = Path.Combine(Path.GetTempPath(), "OpenXmlHtmlTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(dir);
        try
        {
            var htmlPath = Path.Combine(dir, "page.html");
            var docxPath = Path.Combine(dir, "page.docx");

            File.WriteAllText(htmlPath,
                """
                <!DOCTYPE html>
                <html>
                <head><title>Test Page</title></head>
                <body>
                <h1>Welcome</h1>
                <p>This is a <a href="https://example.com">link</a>.</p>
                <table>
                <tr><th>Name</th><th>Value</th></tr>
                <tr><td>Alpha</td><td>100</td></tr>
                <tr><td>Beta</td><td>200</td></tr>
                </table>
                <blockquote>A wise quote</blockquote>
                <pre>  code block
                  indented</pre>
                </body>
                </html>
                """);

            WordHtmlConverter.ConvertFileToDocx(htmlPath, docxPath);

            await VerifyFile(docxPath).UniqueForTargetFrameworkAndVersion();
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
