[TestFixture]
public class WordBlockTests
{
    [Test]
    public Task Paragraphs() =>
        Verify(WordHtmlConverter.ToParagraphs("<p>first paragraph</p><p>second paragraph</p>"));

    [Test]
    public Task Divs() =>
        Verify(WordHtmlConverter.ToParagraphs("<div>first div</div><div>second div</div>"));

    [Test]
    public Task Headings() =>
        Verify(WordHtmlConverter.ToParagraphs("<h1>heading one</h1><h2>heading two</h2>"));

    [Test]
    public Task MixedBlocksAndInline() =>
        Verify(WordHtmlConverter.ToParagraphs("<p>text with <b>bold</b></p><p>another <i>paragraph</i></p>"));

    [Test]
    public Task Blockquote() =>
        Verify(WordHtmlConverter.ToParagraphs("<blockquote>quoted text</blockquote>"));

    [Test]
    public Task PreformattedText() =>
        Verify(WordHtmlConverter.ToParagraphs("<pre>  preserved\n  whitespace</pre>"));

    [Test]
    public Task HorizontalRule() =>
        Verify(WordHtmlConverter.ToParagraphs("above<hr>below"));

    [Test]
    public Task UnorderedList() =>
        Verify(WordHtmlConverter.ToParagraphs("<ul><li>first</li><li>second</li></ul>"));

    [Test]
    public Task OrderedList() =>
        Verify(WordHtmlConverter.ToParagraphs("<ol><li>first</li><li>second</li><li>third</li></ol>"));

    [Test]
    public Task FormattedListItems() =>
        Verify(WordHtmlConverter.ToParagraphs("<ul><li><b>bold</b> item</li><li><i>italic</i> item</li></ul>"));

    [Test]
    public Task NestedUnorderedLists() =>
        Verify(WordHtmlConverter.ToParagraphs("<ul><li>outer</li><li><ul><li>inner</li></ul></li></ul>"));

    [Test]
    public Task NestedOrderedList() =>
        Verify(WordHtmlConverter.ToParagraphs("<ol><li>first</li><li><ol><li>nested first</li><li>nested second</li></ol></li><li>second</li></ol>"));

    [Test]
    public Task DeeplyNestedList() =>
        Verify(WordHtmlConverter.ToParagraphs("<ul><li>level 0</li><li><ul><li>level 1</li><li><ul><li>level 2</li></ul></li></ul></li></ul>"));

    [Test]
    public Task MixedNestedLists() =>
        Verify(WordHtmlConverter.ToParagraphs("<ul><li>bullet</li><li><ol><li>numbered</li></ol></li></ul>"));
}
