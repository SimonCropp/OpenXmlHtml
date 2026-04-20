[TestFixture]
public class WordImageSizingTests
{
    const string png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";

    static List<OpenXmlElement> Build(string html)
    {
        var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        return WordHtmlConverter.ToElements(html, main);
    }

    [Test]
    public Task CssWidthPx() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="width: 400px"></p>"""));

    [Test]
    public Task CssHeightPx() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="height: 200px"></p>"""));

    [Test]
    public Task CssWidthAndHeightPx() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="width: 400px; height: 200px"></p>"""));

    [Test]
    public Task CssWidthInches() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="width: 2in"></p>"""));

    [Test]
    public Task CssWidthEm() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="width: 10em; height: 5em"></p>"""));

    [Test]
    public Task CssWidthPt() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="width: 300pt"></p>"""));

    [Test]
    public Task CssOverridesHtmlAttribute() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" width="50" height="50" style="width: 400px; height: 200px"></p>"""));

    [Test]
    public Task HtmlAttributeStillWorks() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" width="75" height="60"></p>"""));

    [Test]
    public Task CssPercentageIgnored() =>
        Verify(Build($"""<p><img src="data:image/png;base64,{png}" style="width: 50%" width="150"></p>"""));
}
