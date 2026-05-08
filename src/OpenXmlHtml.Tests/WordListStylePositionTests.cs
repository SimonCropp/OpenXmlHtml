[TestFixture]
public class WordListStylePositionTests
{
    [Test]
    public Task UnorderedInside()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ul style="list-style-position: inside">
              <li>First</li>
              <li>Second</li>
            </ul>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task OrderedInside()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol style="list-style-position: inside">
              <li>First</li>
              <li>Second</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ExplicitOutside()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ul style="list-style-position: outside">
              <li>First</li>
              <li>Second</li>
            </ul>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
