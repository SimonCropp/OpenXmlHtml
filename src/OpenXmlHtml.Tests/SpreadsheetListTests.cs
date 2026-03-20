[TestFixture]
public class SpreadsheetListTests
{
    [Test]
    public Task UnorderedList() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ul><li>item one</li><li>item two</li><li>item three</li></ul>"));

    [Test]
    public Task OrderedList() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ol><li>first</li><li>second</li><li>third</li></ol>"));

    [Test]
    public Task SingleListItem() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ul><li>only item</li></ul>"));

    [Test]
    public Task FormattedListItems() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ul><li><b>bold item</b></li><li><i>italic item</i></li></ul>"));

    [Test]
    public Task NestedLists() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ul><li>outer</li><li><ul><li>inner</li></ul></li></ul>"));

    [Test]
    public Task DeeplyNestedLists() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ul><li>level 0</li><li><ul><li>level 1</li><li><ul><li>level 2</li></ul></li></ul></li></ul>"));

    [Test]
    public Task NestedOrderedList() =>
        Verify(SpreadsheetHtmlConverter.ToInlineString("<ol><li>first</li><li><ol><li>nested</li></ol></li></ol>"));
}
