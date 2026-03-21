#if NETFRAMEWORK
using System.Net.Http;
#endif

[TestFixture]
public class WordRemoteImageTests
{
    // 2x2 red PNG
    static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==");

    static HttpClient CreateFakeClient() =>
        new(new FakeImageHandler());

    [Test]
    public Task WebImage_AllowAll()
    {
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.AllowAll(),
            HttpClient = CreateFakeClient()
        };

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><img src="https://example.com/logo.png" width="50" height="50"></p>""",
            stream,
            settings);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task WebImage_SafeDomains_Allowed()
    {
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.SafeDomains("example.com"),
            HttpClient = CreateFakeClient()
        };

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><img src="https://cdn.example.com/logo.png" width="50" height="50"></p>""",
            stream,
            settings);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task WebImage_SafeDomains_Denied()
    {
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.SafeDomains("trusted.com"),
            HttpClient = CreateFakeClient()
        };

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><img src="https://evil.com/logo.png" alt="Blocked Image"></p>""",
            stream,
            settings);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task WebImage_Filter()
    {
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.Filter(src => src.Contains("allowed")),
            HttpClient = CreateFakeClient()
        };

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><img src="https://allowed.example.com/logo.png" width="50" height="50"></p>""",
            stream,
            settings);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public async Task LocalImage_SafeDirectory()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "OpenXmlHtmlTest_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);
        try
        {
            var imagePath = Path.Combine(tempDir, "test.png");
            File.WriteAllBytes(imagePath, PngBytes);

            var settings = new HtmlConvertSettings
            {
                LocalImages = ImagePolicy.SafeDirectories(tempDir)
            };

            using var stream = new MemoryStream();
            WordHtmlConverter.ConvertToDocx(
                $"""<p><img src="{imagePath.Replace("\\", "/")}" width="50" height="50"></p>""",
                stream,
                settings);
            stream.Position = 0;
            await Verify(stream, "docx");
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public Task LocalImage_Denied_FallsBackToAlt()
    {
        var settings = new HtmlConvertSettings
        {
            LocalImages = ImagePolicy.Deny()
        };

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p><img src="C:\Images\photo.png" alt="Photo"></p>""",
            stream,
            settings);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    class FakeImageHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            var content = new ByteArrayContent(PngBytes);
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/png");
            var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = content
            };
            return Task.FromResult(response);
        }
    }
}
