[TestFixture]
public class WordListBulletGlyphTests
{
    static AbstractNum GetBulletAbstractNum(MainDocumentPart main) =>
        main.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>()
            .Single(a =>
                a.Elements<Level>().FirstOrDefault(_ => _.LevelIndex?.Value == 0) is { } l &&
                l.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet);

    static (string glyph, string font) ReadLevel(AbstractNum abs, int ilvl)
    {
        var level = abs.Elements<Level>().Single(_ => _.LevelIndex?.Value == ilvl);
        var glyph = level.LevelText!.Val!.Value!;
        var fonts = level.NumberingSymbolRunProperties!.GetFirstChild<RunFonts>()!;
        return (glyph, fonts.Ascii!.Value!);
    }

    [Test]
    public void BulletLevelsUseFontGlyphsNotUnicodeBullets()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        WordHtmlConverter.ToElements("<ul><li>a</li></ul>", main);

        var abs = GetBulletAbstractNum(main);

        Assert.Multiple(() =>
        {
            Assert.That(ReadLevel(abs, 0), Is.EqualTo(("\uF0B7", "Symbol")));
            Assert.That(ReadLevel(abs, 1), Is.EqualTo(("o", "Courier New")));
            Assert.That(ReadLevel(abs, 2), Is.EqualTo(("\uF0A7", "Wingdings")));
            Assert.That(ReadLevel(abs, 3), Is.EqualTo(("\uF0B7", "Symbol")));
            Assert.That(ReadLevel(abs, 4), Is.EqualTo(("o", "Courier New")));
            Assert.That(ReadLevel(abs, 5), Is.EqualTo(("\uF0A7", "Wingdings")));
        });
    }

    [Test]
    public void ListParagraphsHaveContextualSpacing()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        var elements = WordHtmlConverter.ToElements(
            """
            <ul>
              <li>Bullet</li>
            </ul>
            <ol>
              <li>Numbered</li>
            </ol>
            """,
            main);

        var listParagraphs = elements
            .OfType<Paragraph>()
            .Where(p => p.ParagraphProperties?.GetFirstChild<NumberingProperties>() != null)
            .ToList();

        Assert.That(listParagraphs, Has.Count.EqualTo(2));
        Assert.Multiple(() =>
        {
            foreach (var p in listParagraphs)
            {
                Assert.That(
                    p.ParagraphProperties!.GetFirstChild<ContextualSpacing>(),
                    Is.Not.Null,
                    "list paragraph must set w:contextualSpacing so consecutive list items render tight");
            }
        });
    }
}
