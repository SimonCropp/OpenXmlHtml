[TestFixture]
public class WordHeadingTests
{
    [Test]
    public Task H1() =>
        Verify(WordHtmlConverter.ToParagraphs("<h1>Main Title</h1>"));

    [Test]
    public Task H2() =>
        Verify(WordHtmlConverter.ToParagraphs("<h2>Subtitle</h2>"));

    [Test]
    public Task H3() =>
        Verify(WordHtmlConverter.ToParagraphs("<h3>Section</h3>"));

    [Test]
    public Task HeadingWithInlineFormatting() =>
        Verify(WordHtmlConverter.ToParagraphs("<h1>Title with <i>italic</i> word</h1>"));

    [Test]
    public Task HeadingFollowedByParagraph() =>
        Verify(WordHtmlConverter.ToParagraphs("<h2>Heading</h2><p>Body text</p>"));

    [Test]
    public Task HeadingStyles() =>
        Verify(WordHtmlConverter.ToElements(
            """
            <h1>Heading 1</h1>
            <h2>Heading 2</h2>
            <h3>Heading 3</h3>
            <h4>Heading 4</h4>
            <h5>Heading 5</h5>
            <h6>Heading 6</h6>
            <p>Normal paragraph</p>
            """));

    [Test]
    public Task HeadingOffsetShifts() =>
        Verify(
            WordHtmlConverter.ToElements(
                """
                <h1>Heading 1</h1>
                <h2>Heading 2</h2>
                <h3>Heading 3</h3>
                """,
                main: null,
                new()
                {
                    HeadingLevelOffset = 1
                }));

    [Test]
    public Task HeadingOffsetClampsAtNine() =>
        Verify(WordHtmlConverter.ToElements(
            "<h5>Deep</h5><h6>Deeper</h6>",
            main: null,
            new()
            {
                HeadingLevelOffset = 5
            }));

    [Test]
    public Task HeadingStylesDocx()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <h1>Chapter One</h1>
            <p>Introduction text.</p>
            <h2>Section 1.1</h2>
            <p>Details here.</p>
            <h2>Section 1.2</h2>
            <p>More details.</p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
