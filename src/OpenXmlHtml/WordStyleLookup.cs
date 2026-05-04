enum StyleType
{
    Paragraph,
    Character
}

static class WordStyleLookup
{
    internal static Dictionary<string, StyleType>? BuildStyleMap(MainDocumentPart? main)
    {
        var styles = main?.StyleDefinitionsPart;
        if (styles?.Styles == null)
        {
            return null;
        }

        var map = new Dictionary<string, StyleType>(StringComparer.OrdinalIgnoreCase);
        foreach (var style in styles.Styles.Elements<Style>())
        {
            var styleId = style.StyleId?.Value;
            if (styleId == null)
            {
                continue;
            }

            if (style.Type?.Value == StyleValues.Paragraph)
            {
                map[styleId] = StyleType.Paragraph;
            }
            else if (style.Type?.Value == StyleValues.Character)
            {
                map[styleId] = StyleType.Character;
            }
        }

        return map.Count > 0 ? map : null;
    }

    internal static (string? ParagraphStyleId, string? RunStyleId) LookupClasses(
        IElement element,
        Dictionary<string, StyleType> styleMap)
    {
        string? paragraphStyleId = null;
        string? runStyleId = null;

        foreach (var className in element.ClassList)
        {
            if (!styleMap.TryGetValue(className, out var type))
            {
                continue;
            }

            switch (type)
            {
                case StyleType.Paragraph when paragraphStyleId == null:
                    paragraphStyleId = className;
                    break;
                case StyleType.Character when runStyleId == null:
                    runStyleId = className;
                    break;
            }

            if (paragraphStyleId != null && runStyleId != null)
            {
                break;
            }
        }

        return (paragraphStyleId, runStyleId);
    }
}
