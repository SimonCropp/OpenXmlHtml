static partial class WordContentBuilder
{
    static void BuildList(IElement element, string tag, FormatState newFormat, List<OpenXmlElement> elements, WordBuildContext context, bool inPre)
    {
        FlushParagraph(elements, context);
        var isReversed = tag == "ol" && element.HasAttribute("reversed");

        if (context.MainPart != null && !isReversed)
        {
            BuildRealNumberedList(element, tag, newFormat, elements, context, inPre);
        }
        else
        {
            if (isReversed)
            {
                var liCount = 0;
                foreach (var child in element.Children)
                {
                    if (child.LocalName == "li")
                    {
                        liCount++;
                    }
                }

                var startAttr2 = element.GetAttribute("start");
                context.ReversedStart = startAttr2 != null && int.TryParse(startAttr2, out var s) ? s : liCount;
            }

            ProcessChildren(element, newFormat, elements, context, inPre);
            context.ReversedStart = null;
        }

        FlushParagraph(elements, context);
    }

    static void BuildRealNumberedList(IElement element, string tag, FormatState newFormat, List<OpenXmlElement> elements, WordBuildContext context, bool inPre)
    {
        var isOrdered = tag == "ol";
        var part = WordNumberingBuilder.EnsureNumberingPart(context.MainPart!);
        var numbering = part.Numbering!;

        var typeAttr = element.GetAttribute("type");
        var listStyleCss = element.GetAttribute("style") is { } listStyle
            ? StyleParser.Parse(listStyle).GetValueOrDefault("list-style-type")
            : null;
        var format = WordNumberingBuilder.ParseListStyleType(typeAttr, listStyleCss, isOrdered);

        int abstractNumId;
        if (format == NumberFormatValues.Bullet)
        {
            if (context.BulletAbstractNumId == null)
            {
                var id = WordNumberingBuilder.GetNextId(numbering);
                context.BulletAbstractNumId = WordNumberingBuilder.CreateBulletAbstractNum(numbering, id);
            }

            abstractNumId = context.BulletAbstractNumId.Value;
        }
        else
        {
            // Each ordered format type gets its own abstract num
            var id = WordNumberingBuilder.GetNextId(numbering);
            abstractNumId = WordNumberingBuilder.CreateOrderedAbstractNum(numbering, id, format);
        }

        int? startOverride = null;
        var startAttr = element.GetAttribute("start");
        if (startAttr != null && int.TryParse(startAttr, out var startVal) && startVal != 1)
        {
            startOverride = startVal;
        }

        var numId = WordNumberingBuilder.GetNextId(numbering);
        WordNumberingBuilder.AddNumberingInstance(numbering, numId, abstractNumId, startOverride);
        var ilvl = context.ListStack.Count > 0 ? context.ListStack.Peek().Ilvl + 1 : 0;
        context.ListStack.Push((numId, ilvl, isOrdered));
        ProcessChildren(element, newFormat, elements, context, inPre);
        context.ListStack.Pop();
    }

    static void BuildListItem(IElement element, FormatState newFormat, List<OpenXmlElement> elements, WordBuildContext context, bool inPre)
    {
        FlushParagraph(elements, context);
        if (context.ListStack.Count > 0)
        {
            var (numId, ilvl, _) = context.ListStack.Peek();
            context.ListNumId = numId;
            context.ListIlvl = ilvl;
        }
        else
        {
            // Fallback: text prefix when no MainDocumentPart
            var parent = element.ParentElement?.LocalName;
            var depth = HtmlSegmentParser.GetListDepth(element);
            context.ListDepth = depth;

            if (parent == "ol")
            {
                var index = 1;
                foreach (var sibling in element.ParentElement!.Children)
                {
                    if (sibling == element)
                    {
                        break;
                    }

                    if (sibling.LocalName == "li")
                    {
                        index++;
                    }
                }

                if (context.ReversedStart != null)
                {
                    index = context.ReversedStart.Value - (index - 1);
                }

                AddTextRun($"{index}. ", newFormat, context);
            }
            else
            {
                var bullet = depth switch
                {
                    0 => "\u25CF",
                    1 => "\u25CB",
                    _ => "\u25A0"
                };
                AddTextRun($"{bullet} ", newFormat, context);
            }
        }

        ProcessChildren(element, newFormat, elements, context, inPre);
        FlushParagraph(elements, context);
    }
}
