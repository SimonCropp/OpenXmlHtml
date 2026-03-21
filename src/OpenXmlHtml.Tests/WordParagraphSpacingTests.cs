[TestFixture]
public class WordParagraphSpacingTests
{
    [Test]
    public Task MarginTop()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="margin-top: 24pt">Spaced paragraph</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MarginBottom()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="margin-bottom: 12pt">Bottom margin</p><p>Next paragraph</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MarginLeftRight()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="margin-left: 36pt; margin-right: 36pt">Indented paragraph</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MarginShorthand()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="margin: 12pt 24pt">Shorthand margins</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TextIndent()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="text-indent: 36pt">First line indented</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TextAlignCenter()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="text-align: center">Centered text</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task TextAlignJustify()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="text-align: justify">Justified text with enough content to show justification effect clearly in the document.</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task LineHeightMultiple()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="line-height: 2">Double spaced line</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task LineHeightPercent()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="line-height: 150%">One and a half spacing</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task LineHeightFixed()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """<p style="line-height: 18pt">Fixed 18pt line height</p>""",
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task CombinedStyles()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <p style="text-align: center; margin-top: 24pt; margin-bottom: 12pt">Centered with spacing</p>
            <p style="margin-left: 72pt; text-indent: 36pt; line-height: 1.5">Indented with first line and 1.5 spacing</p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
