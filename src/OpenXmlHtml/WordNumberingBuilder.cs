static class WordNumberingBuilder
{
    static readonly string[] levelIndents = ["720", "1440", "2160", "2880", "3600", "4320", "5040", "5760", "6480"];

    internal static NumberingDefinitionsPart EnsureNumberingPart(MainDocumentPart main)
    {
        var part = main.NumberingDefinitionsPart;
        if (part != null)
        {
            return part;
        }

        part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new();
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
        var abstractNum = new AbstractNum
        {
            AbstractNumberId = abstractNumId
        };
        abstractNum.Append(new MultiLevelType
        {
            Val = MultiLevelValues.HybridMultilevel
        });

        var bullets = new[]
        {
            "\u25CF",
            "\u25CB",
            "\u25A0"
        };
        for (var i = 0; i < 9; i++)
        {
            var bullet = bullets[i % bullets.Length];
            var level = new Level(
                new StartNumberingValue
                {
                    Val = 1
                },
                new NumberingFormat
                {
                    Val = NumberFormatValues.Bullet
                },
                new LevelText
                {
                    Val = bullet
                },
                new LevelJustification
                {
                    Val = LevelJustificationValues.Left
                },
                new ParagraphProperties(
                    new Indentation
                    {
                        Left = levelIndents[i],
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

        InsertAbstractNum(numbering, abstractNum);
        return abstractNumId;
    }

    internal static int CreateOrderedAbstractNum(Numbering numbering, int abstractNumId, NumberFormatValues format)
    {
        var abstractNum = new AbstractNum
        {
            AbstractNumberId = abstractNumId
        };
        abstractNum.Append(new MultiLevelType
        {
            Val = MultiLevelValues.HybridMultilevel
        });

        var suffix = format == NumberFormatValues.Decimal ? "." : ")";
        for (var i = 0; i < 9; i++)
        {
            var level = new Level(
                new StartNumberingValue
                {
                    Val = 1
                },
                new NumberingFormat
                {
                    Val = format
                },
                new LevelText
                {
                    Val = $"%{i + 1}{suffix}"
                },
                new LevelJustification
                {
                    Val = LevelJustificationValues.Left
                },
                new ParagraphProperties(
                    new Indentation
                    {
                        Left = levelIndents[i],
                        Hanging = "360"
                    }))
            {
                LevelIndex = i
            };
            abstractNum.Append(level);
        }

        InsertAbstractNum(numbering, abstractNum);
        return abstractNumId;
    }

    internal static int AddNumberingInstance(Numbering numbering, int numId, int abstractNumId, int? startOverride = null)
    {
        var instance = new NumberingInstance(
            new AbstractNumId
            {
                Val = abstractNumId
            })
        {
            NumberID = numId
        };

        if (startOverride != null)
        {
            instance.Append(
                new LevelOverride(
                    new StartOverrideNumberingValue
                    {
                        Val = startOverride.Value
                    })
                {
                    LevelIndex = 0
                });
        }

        numbering.Append(instance);
        return numId;
    }

    internal static NumberFormatValues ParseListStyleType(string? type, string? cssListStyle, bool isOrdered)
    {
        var val = (cssListStyle ?? type).AsSpan().Trim();
        if (val.Equals("a", StringComparison.Ordinal) ||
            val.Equals("lower-alpha", StringComparison.OrdinalIgnoreCase) ||
            val.Equals("lower-latin", StringComparison.OrdinalIgnoreCase))
        {
            return NumberFormatValues.LowerLetter;
        }

        if (val.Equals("A", StringComparison.Ordinal) ||
            val.Equals("upper-alpha", StringComparison.OrdinalIgnoreCase) ||
            val.Equals("upper-latin", StringComparison.OrdinalIgnoreCase))
        {
            return NumberFormatValues.UpperLetter;
        }

        if (val.Equals("i", StringComparison.Ordinal) ||
            val.Equals("lower-roman", StringComparison.OrdinalIgnoreCase))
        {
            return NumberFormatValues.LowerRoman;
        }

        if (val.Equals("I", StringComparison.Ordinal) ||
            val.Equals("upper-roman", StringComparison.OrdinalIgnoreCase))
        {
            return NumberFormatValues.UpperRoman;
        }

        if (val.Equals("1", StringComparison.Ordinal) ||
            val.Equals("decimal", StringComparison.OrdinalIgnoreCase))
        {
            return NumberFormatValues.Decimal;
        }

        return isOrdered ? NumberFormatValues.Decimal : NumberFormatValues.Bullet;
    }

    static void InsertAbstractNum(Numbering numbering, AbstractNum abstractNum)
    {
        var firstInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstInstance != null)
        {
            numbering.InsertBefore(abstractNum, firstInstance);
        }
        else
        {
            numbering.Append(abstractNum);
        }
    }
}
