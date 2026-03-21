[TestFixture]
public class WordReversedListTests
{
    [Test]
    public Task ReversedOrderedList()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol reversed>
              <li>Third (should be 3)</li>
              <li>Second (should be 2)</li>
              <li>First (should be 1)</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ReversedWithStart()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol reversed start="10">
              <li>Should be 10</li>
              <li>Should be 9</li>
              <li>Should be 8</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ReversedFallbackWithoutMainPart()
    {
        var elements = WordHtmlConverter.ToElements(
            """
            <ol reversed>
              <li>Third</li>
              <li>Second</li>
              <li>First</li>
            </ol>
            """);
        return Verify(elements);
    }
}
