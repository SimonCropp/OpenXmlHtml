[TestFixture]
public class StyleParserTests
{
    [Test]
    public void ParseSingleProperty()
    {
        var result = StyleParser.Parse("color: red");
        Assert.That(result["color"], Is.EqualTo("red"));
    }

    [Test]
    public void ParseMultipleProperties()
    {
        var result = StyleParser.Parse("font-weight: bold; font-style: italic; color: blue");
        Assert.That(result["font-weight"], Is.EqualTo("bold"));
        Assert.That(result["font-style"], Is.EqualTo("italic"));
        Assert.That(result["color"], Is.EqualTo("blue"));
    }

    [Test]
    public void ParseNullStyle()
    {
        var result = StyleParser.Parse(null);
        Assert.That(result, Is.Empty);
    }

    [Test]
    public void ParseEmptyStyle()
    {
        var result = StyleParser.Parse("");
        Assert.That(result, Is.Empty);
    }

    [Test]
    public void ParseTrailingSemicolon()
    {
        var result = StyleParser.Parse("color: red;");
        Assert.That(result["color"], Is.EqualTo("red"));
    }

    [Test]
    public void CaseInsensitiveKeys()
    {
        var result = StyleParser.Parse("Color: red");
        Assert.That(result["color"], Is.EqualTo("red"));
    }

    [Test]
    public void FontSizePt() =>
        Assert.That(StyleParser.ParseFontSize("12pt"), Is.EqualTo(12));

    [Test]
    public void FontSizePx() =>
        Assert.That(StyleParser.ParseFontSize("16px"), Is.EqualTo(12));

    [Test]
    public void FontSizeEm() =>
        Assert.That(StyleParser.ParseFontSize("2em"), Is.EqualTo(24));

    [Test]
    public void FontSizeKeywords()
    {
        Assert.That(StyleParser.ParseFontSize("xx-small"), Is.EqualTo(7));
        Assert.That(StyleParser.ParseFontSize("x-small"), Is.EqualTo(8));
        Assert.That(StyleParser.ParseFontSize("small"), Is.EqualTo(10));
        Assert.That(StyleParser.ParseFontSize("medium"), Is.EqualTo(12));
        Assert.That(StyleParser.ParseFontSize("large"), Is.EqualTo(14));
        Assert.That(StyleParser.ParseFontSize("x-large"), Is.EqualTo(18));
        Assert.That(StyleParser.ParseFontSize("xx-large"), Is.EqualTo(24));
    }

    [Test]
    public void FontSizeRawNumber() =>
        Assert.That(StyleParser.ParseFontSize("14"), Is.EqualTo(14));

    [Test]
    public void FontSizeInvalid() =>
        Assert.That(StyleParser.ParseFontSize("abc"), Is.Null);
}
