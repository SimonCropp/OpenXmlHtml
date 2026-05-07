[TestFixture]
public class WordListNumberingTests
{
    [Test]
    public Task UnorderedList()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ul>
              <li>Alpha</li>
              <li>Beta</li>
              <li>Gamma</li>
            </ul>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task OrderedList()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol>
              <li>First</li>
              <li>Second</li>
              <li>Third</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task NestedList()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <p>Intro paragraph</p>
            <ul>
              <li>Top level
                <ul>
                  <li>Nested item</li>
                  <li>Another nested</li>
                </ul>
              </li>
              <li>Back to top</li>
            </ul>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MixedOrderedAndUnordered()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol>
              <li>Numbered
                <ul>
                  <li>Bulleted child</li>
                </ul>
              </li>
              <li>Another numbered</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task SeparateListsRestartNumbering()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol>
              <li>First list item 1</li>
              <li>First list item 2</li>
            </ol>
            <p>Paragraph between lists</p>
            <ol>
              <li>Second list item 1</li>
              <li>Second list item 2</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task FallbackWithoutMainDocumentPart()
    {
        var elements = WordHtmlConverter.ToElements(
            """
            <ul>
              <li>Bullet with text prefix</li>
              <li>Another item</li>
            </ul>
            """);
        return Verify(elements);
    }
}
