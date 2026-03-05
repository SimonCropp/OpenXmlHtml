[TestFixture]
public class WordNestedTests
{
    [Test]
    public Task BoldItalic() =>
        Verify(WordHtmlConverter.ToParagraphs("<b><i>bold italic</i></b>"));

    [Test]
    public Task BoldUnderlineItalic() =>
        Verify(WordHtmlConverter.ToParagraphs("<b><u><i>all three</i></u></b>"));

    [Test]
    public Task PartialOverlap() =>
        Verify(WordHtmlConverter.ToParagraphs("<b>bold <i>bold-italic</i> bold</b>"));

    [Test]
    public Task DeeplyNested() =>
        Verify(WordHtmlConverter.ToParagraphs("<b><i><u><s>all formats</s></u></i></b>"));

    [Test]
    public Task MixedContent() =>
        Verify(WordHtmlConverter.ToParagraphs("start <b>bold <i>both</i></b> <u>under</u> end"));

    [Test]
    public Task NestedColors() =>
        Verify(WordHtmlConverter.ToParagraphs(
            "<span style=\"color: red\">outer <span style=\"color: blue\">inner</span> outer</span>"));
}
