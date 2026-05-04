#if NETFRAMEWORK
using System.Net.Http;
#endif

[TestFixture]
public class WordImageFallbackTests
{
    [Test]
    public void InvalidBase64DataUri_FallsBackToAlt()
    {
        var elements = WordHtmlConverter.ToElements(
            """<p><img src="data:image/png;base64,!!!notbase64!!!" alt="Bad"></p>""");

        AssertNoDrawing(elements);
        AssertContainsText(elements, "Bad");
    }

    [Test]
    public void NonBase64DataUri_FallsBackToAlt()
    {
        var elements = WordHtmlConverter.ToElements(
            """<p><img src="data:image/png,raw" alt="Bad"></p>""");

        AssertNoDrawing(elements);
        AssertContainsText(elements, "Bad");
    }

    [Test]
    public void ImageWithoutMainPart_SilentlyDropped()
    {
        var png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";
        var elements = WordHtmlConverter.ToElements(
            $"""<p><img src="data:image/png;base64,{png}"></p>""");

        AssertNoDrawing(elements);
    }

    [Test]
    public void HttpThrows_FallsBackToAlt()
    {
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.AllowAll(),
            HttpClient = new(new ThrowingHandler())
        };

        var elements = WordHtmlConverter.ToElements(
            """<p><img src="https://example.com/img.png" alt="NoNet"></p>""",
            null,
            settings);

        AssertNoDrawing(elements);
        AssertContainsText(elements, "NoNet");
    }

    [Test]
    public void HttpNotSuccess_FallsBackToAlt()
    {
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.AllowAll(),
            HttpClient = new(new StatusCodeHandler(System.Net.HttpStatusCode.InternalServerError))
        };

        var elements = WordHtmlConverter.ToElements(
            """<p><img src="https://example.com/img.png" alt="500"></p>""",
            null,
            settings);

        AssertNoDrawing(elements);
        AssertContainsText(elements, "500");
    }

    [Test]
    public void LocalImageMissing_FallsBackToAlt()
    {
        var missingPath = Path.Combine(Path.GetTempPath(), "definitely_missing_" + Guid.NewGuid().ToString("N") + ".png");
        var settings = new HtmlConvertSettings
        {
            LocalImages = ImagePolicy.AllowAll()
        };

        var elements = WordHtmlConverter.ToElements(
            $"""<p><img src="{missingPath.Replace("\\", "/")}" alt="Missing"></p>""",
            null,
            settings);

        AssertNoDrawing(elements);
        AssertContainsText(elements, "Missing");
    }

    [Test]
    public void LocalImageInvalidPath_FallsBackToAlt()
    {
        var settings = new HtmlConvertSettings
        {
            LocalImages = ImagePolicy.AllowAll()
        };

        var elements = WordHtmlConverter.ToElements(
            """<p><img src="\0invalid" alt="Invalid"></p>""",
            null,
            settings);

        AssertNoDrawing(elements);
        AssertContainsText(elements, "Invalid");
    }

    [Test]
    public void LocalImageMalformedFileUri_FallsBackToAlt()
    {
        var settings = new HtmlConvertSettings
        {
            LocalImages = ImagePolicy.AllowAll()
        };

        var elements = WordHtmlConverter.ToElements(
            """<p><img src="file:///%" alt="Malformed"></p>""",
            null,
            settings);

        AssertNoDrawing(elements);
        AssertContainsText(elements, "Malformed");
    }

    static void AssertNoDrawing(List<OpenXmlElement> elements)
    {
        var hasDrawing = elements.Any(e => e.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
        Assert.That(hasDrawing, Is.False, "Expected no Drawing elements");
    }

    static void AssertContainsText(List<OpenXmlElement> elements, string text)
    {
        var combined = string.Concat(elements.SelectMany(e => e.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)));
        Assert.That(combined, Does.Contain(text));
    }

    class ThrowingHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Cancel cancel) =>
            throw new HttpRequestException("boom");
    }

    class StatusCodeHandler(System.Net.HttpStatusCode code) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Cancel cancel) =>
            Task.FromResult(new HttpResponseMessage(code));
    }
}
