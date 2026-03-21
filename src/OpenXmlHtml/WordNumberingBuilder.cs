namespace OpenXmlHtml;

static class WordNumberingBuilder
{
    internal static NumberingDefinitionsPart EnsureNumberingPart(MainDocumentPart mainPart)
    {
        var part = mainPart.NumberingDefinitionsPart;
        if (part != null)
        {
            return part;
        }

        part = mainPart.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new Numbering();
        return part;
    }

    internal static int GetNextId(Numbering numbering)
    {
        var maxId = 0;
        foreach (var abs in numbering.Elements<AbstractNum>())
        {
            if (abs.AbstractNumberId?.Value is { } id && id > maxId)
            {
                maxId = id;
            }
        }

        foreach (var num in numbering.Elements<NumberingInstance>())
        {
            if (num.NumberID?.Value is { } id && id > maxId)
            {
                maxId = id;
            }
        }

        return maxId + 1;
    }

    internal static int CreateBulletAbstractNum(Numbering numbering, int abstractNumId)
    {
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        var bullets = new[] { "\u25CF", "\u25CB", "\u25A0" };
        for (var i = 0; i < 9; i++)
        {
            var bullet = bullets[i % bullets.Length];
            var level = new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = bullet },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Indentation
                    {
                        Left = ((i + 1) * 720).ToString(),
                        Hanging = "360"
                    }),
                new NumberingSymbolRunProperties(
                    new RunFonts
                    {
                        Ascii = "Symbol",
                        HighAnsi = "Symbol"
                    }))
            {
                LevelIndex = i
            };
            abstractNum.Append(level);
        }

        // Insert before any NumberingInstance elements
        var firstInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstInstance != null)
        {
            numbering.InsertBefore(abstractNum, firstInstance);
        }
        else
        {
            numbering.Append(abstractNum);
        }

        return abstractNumId;
    }

    internal static int CreateDecimalAbstractNum(Numbering numbering, int abstractNumId)
    {
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        for (var i = 0; i < 9; i++)
        {
            var level = new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = $"%{i + 1}." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Indentation
                    {
                        Left = ((i + 1) * 720).ToString(),
                        Hanging = "360"
                    }))
            {
                LevelIndex = i
            };
            abstractNum.Append(level);
        }

        var firstInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstInstance != null)
        {
            numbering.InsertBefore(abstractNum, firstInstance);
        }
        else
        {
            numbering.Append(abstractNum);
        }

        return abstractNumId;
    }

    internal static int AddNumberingInstance(Numbering numbering, int numId, int abstractNumId)
    {
        numbering.Append(
            new NumberingInstance(
                new AbstractNumId { Val = abstractNumId })
            {
                NumberID = numId
            });
        return numId;
    }
}
