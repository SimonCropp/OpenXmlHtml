[TestFixture]
public class WordListStyleTests
{
    [Test]
    public void UnknownListStyleType_OrderedFallsBackToDecimal() =>
        Assert.That(
            WordNumberingBuilder.ParseListStyleType("nonsense", null, isOrdered: true),
            Is.EqualTo(NumberFormatValues.Decimal));

    [Test]
    public void UnknownListStyleType_UnorderedFallsBackToBullet() =>
        Assert.That(
            WordNumberingBuilder.ParseListStyleType(null, "nonsense", isOrdered: false),
            Is.EqualTo(NumberFormatValues.Bullet));

    [Test]
    public Task LowerAlphaType()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol type="a">
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
    public Task UpperAlphaType()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol type="A">
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
    public Task LowerRomanType()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol type="i">
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
    public Task UpperRomanType()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol type="I">
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
    public Task CssListStyleType()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol style="list-style-type: lower-roman">
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
    public Task StartAttribute()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol start="5">
              <li>Should be 5</li>
              <li>Should be 6</li>
              <li>Should be 7</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task StartAttributeWithAlpha()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol type="a" start="3">
              <li>Should be c</li>
              <li>Should be d</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task MixedListTypes()
    {
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <ol type="I">
              <li>Roman numeral
                <ol type="a">
                  <li>Alpha nested</li>
                  <li>Alpha nested</li>
                </ol>
              </li>
              <li>Roman numeral</li>
            </ol>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
