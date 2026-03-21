[TestFixture]
public class ColorParserTests
{
    [TestCase("#FF0000", "FF0000")]
    [TestCase("#ff0000", "FF0000")]
    [TestCase("#F00", "FF0000")]
    [TestCase("#f00", "FF0000")]
    [TestCase("#12345678", "123456")]
    [TestCase("red", "FF0000")]
    [TestCase("Blue", "0000FF")]
    [TestCase("GREEN", "008000")]
    [TestCase("rgb(255, 0, 0)", "FF0000")]
    [TestCase("rgb(0,128,0)", "008000")]
    [TestCase("rgb(0, 0, 255)", "0000FF")]
    [TestCase("rgba(255, 0, 0, 0.5)", "FF0000")]
    [TestCase("rgba(0, 128, 0, 1)", "008000")]
    [TestCase("rgba(0,0,255,0)", "0000FF")]
    [TestCase("rgba(100, 200, 50, 0.8)", "64C832")]
    public void ValidColors(string input, string expected) =>
        Assert.That(ColorParser.Parse(input), Is.EqualTo(expected));

    [TestCase(null)]
    [TestCase("")]
    [TestCase("   ")]
    [TestCase("notacolor")]
    [TestCase("#GGG")]
    [TestCase("rgb(abc)")]
    public void InvalidColors(string? input) =>
        Assert.That(ColorParser.Parse(input), Is.Null);

    [Test]
    public void NamedColorCoverage()
    {
        Assert.That(ColorParser.Parse("black"), Is.EqualTo("000000"));
        Assert.That(ColorParser.Parse("white"), Is.EqualTo("FFFFFF"));
        Assert.That(ColorParser.Parse("yellow"), Is.EqualTo("FFFF00"));
        Assert.That(ColorParser.Parse("orange"), Is.EqualTo("FFA500"));
        Assert.That(ColorParser.Parse("purple"), Is.EqualTo("800080"));
        Assert.That(ColorParser.Parse("pink"), Is.EqualTo("FFC0CB"));
        Assert.That(ColorParser.Parse("gray"), Is.EqualTo("808080"));
        Assert.That(ColorParser.Parse("grey"), Is.EqualTo("808080"));
        Assert.That(ColorParser.Parse("cyan"), Is.EqualTo("00FFFF"));
        Assert.That(ColorParser.Parse("magenta"), Is.EqualTo("FF00FF"));
        Assert.That(ColorParser.Parse("navy"), Is.EqualTo("000080"));
        Assert.That(ColorParser.Parse("teal"), Is.EqualTo("008080"));
        Assert.That(ColorParser.Parse("maroon"), Is.EqualTo("800000"));
        Assert.That(ColorParser.Parse("olive"), Is.EqualTo("808000"));
        Assert.That(ColorParser.Parse("silver"), Is.EqualTo("C0C0C0"));
        Assert.That(ColorParser.Parse("crimson"), Is.EqualTo("DC143C"));
        Assert.That(ColorParser.Parse("indigo"), Is.EqualTo("4B0082"));
    }

    [Test]
    public void RgbClamping() =>
        Assert.That(ColorParser.Parse("rgb(300, -10, 128)"), Is.EqualTo("FF0080"));
}
