[TestFixture]
public class WordStyleMappingTests
{
    static (MainDocumentPart MainPart, Body Body) CreateDocumentWithStyles(Stream stream, params (string Id, StyleValues Type)[] styles)
    {
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();

        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        var stylesheet = new Styles();
        foreach (var (id, type) in styles)
        {
            stylesheet.Append(
                new Style
                {
                    StyleId = id,
                    Type = type
                });
        }

        stylesPart.Styles = stylesheet;

        var body = new Body();
        main.Document = new(body);
        return (main, body);
    }

    [Test]
    public Task ParagraphStyleFromClass()
    {
        using var stream = new MemoryStream();
        var (main, body) = CreateDocumentWithStyles(stream,
            ("Quote", StyleValues.Paragraph));

        WordHtmlConverter.AppendHtml(body,
            """<p class="Quote">This should use the Quote style</p>""",
            main);

        return Verify(body);
    }

    [Test]
    public Task CharacterStyleFromClass()
    {
        using var stream = new MemoryStream();
        var (main, body) = CreateDocumentWithStyles(stream,
            ("Emphasis", StyleValues.Character));

        WordHtmlConverter.AppendHtml(body,
            """<p>Normal text with <span class="Emphasis">emphasized</span> word</p>""",
            main);

        return Verify(body);
    }

    [Test]
    public Task BothParagraphAndCharacterStyles()
    {
        using var stream = new MemoryStream();
        var (main, body) = CreateDocumentWithStyles(stream,
            ("IntenseQuote", StyleValues.Paragraph),
            ("Strong", StyleValues.Character));

        WordHtmlConverter.AppendHtml(body,
            """<blockquote class="IntenseQuote">Quote with <span class="Strong">strong</span> text</blockquote>""",
            main);

        return Verify(body);
    }

    [Test]
    public Task ClassNotInStyles_NoEffect()
    {
        using var stream = new MemoryStream();
        var (main, body) = CreateDocumentWithStyles(stream,
            ("Quote", StyleValues.Paragraph));

        WordHtmlConverter.AppendHtml(body,
            """<p class="NonExistent">Should render as default</p>""",
            main);

        return Verify(body);
    }

    [Test]
    public Task HeadingStyleTakesPrecedenceOverClass()
    {
        using var stream = new MemoryStream();
        var (main, body) = CreateDocumentWithStyles(stream,
            ("CustomStyle", StyleValues.Paragraph));

        WordHtmlConverter.AppendHtml(body,
            """<h1 class="CustomStyle">Heading should use Heading1 not CustomStyle</h1>""",
            main);

        return Verify(body);
    }

    [Test]
    public Task CaseInsensitiveStyleLookup()
    {
        using var stream = new MemoryStream();
        var (main, body) = CreateDocumentWithStyles(stream,
            ("Quote", StyleValues.Paragraph));

        WordHtmlConverter.AppendHtml(body,
            """<p class="quote">Case insensitive match</p>""",
            main);

        return Verify(body);
    }
}
