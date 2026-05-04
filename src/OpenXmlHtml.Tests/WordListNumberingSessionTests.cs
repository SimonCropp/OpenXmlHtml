[TestFixture]
public class WordListNumberingSessionTests
{
    static int BulletAbstractCount(MainDocumentPart main) =>
        main.NumberingDefinitionsPart?.Numbering?
            .Elements<AbstractNum>()
            .Count(a =>
                a.Elements<Level>().FirstOrDefault(_ => _.LevelIndex?.Value == 0) is { } l &&
                l.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet) ?? 0;

    static int NumberingInstanceCount(MainDocumentPart main) =>
        main.NumberingDefinitionsPart?.Numbering?
            .Elements<NumberingInstance>().Count() ?? 0;

    [Test]
    public void WithoutSession_TwoCallsEachCreateOwnBulletAbstract()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        WordHtmlConverter.ToElements("<ul><li>a</li></ul>", main, new());
        WordHtmlConverter.ToElements("<ul><li>b</li></ul>", main, new());

        Assert.That(BulletAbstractCount(main), Is.EqualTo(2));
        Assert.That(NumberingInstanceCount(main), Is.EqualTo(2));
    }

    [Test]
    public void WithSession_TwoCallsShareOneBulletAbstract()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        var session = new HtmlNumberingSession();
        WordHtmlConverter.ToElements("<ul><li>a</li></ul>", main, new() { NumberingSession = session });
        WordHtmlConverter.ToElements("<ul><li>b</li></ul>", main, new() { NumberingSession = session });

        Assert.That(BulletAbstractCount(main), Is.EqualTo(1));
        Assert.That(NumberingInstanceCount(main), Is.EqualTo(2));
    }

    [Test]
    public void WithSession_SessionPopulatedAfterFirstCall()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        var session = new HtmlNumberingSession();
        Assert.That(session.BulletAbstractNumId, Is.Null);

        WordHtmlConverter.ToElements("<ul><li>a</li></ul>", main, new() { NumberingSession = session });

        Assert.That(session.BulletAbstractNumId, Is.Not.Null);
    }

    [Test]
    public void WithSession_NoBulletList_SessionRemainsNull()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        var session = new HtmlNumberingSession();
        WordHtmlConverter.ToElements("<p>plain text</p>", main, new() { NumberingSession = session });

        Assert.That(session.BulletAbstractNumId, Is.Null);
    }

    [Test]
    public void WithSession_FirstCallNoList_SecondCallBullet_CreatesOneAbstract()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        var session = new HtmlNumberingSession();
        WordHtmlConverter.ToElements("<p>no list here</p>", main, new() { NumberingSession = session });
        WordHtmlConverter.ToElements("<ul><li>a</li><li>b</li></ul>", main, new() { NumberingSession = session });

        Assert.That(BulletAbstractCount(main), Is.EqualTo(1));
        Assert.That(NumberingInstanceCount(main), Is.EqualTo(1));
    }

    [Test]
    public void WithSession_ThreeCallsShareOneBulletAbstract()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(new Body());

        var session = new HtmlNumberingSession();
        WordHtmlConverter.ToElements("<ul><li>a</li></ul>", main, new() { NumberingSession = session });
        WordHtmlConverter.ToElements("<ul><li>b</li></ul>", main, new() { NumberingSession = session });
        WordHtmlConverter.ToElements("<ul><li>c</li></ul>", main, new() { NumberingSession = session });

        Assert.That(BulletAbstractCount(main), Is.EqualTo(1));
        Assert.That(NumberingInstanceCount(main), Is.EqualTo(3));
    }
}
