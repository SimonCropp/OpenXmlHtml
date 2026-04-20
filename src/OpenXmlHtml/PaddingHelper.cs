static class PaddingHelper
{
    internal static (int? Top, int? Right, int? Bottom, int? Left)? TryParse(Dictionary<string, string> declarations)
    {
        int? top = null, right = null, bottom = null, left = null;
        var hasPadding = false;

        if (declarations.TryGetValue("padding", out var shorthand))
        {
            (top, right, bottom, left) = StyleParser.ParseMarginShorthand(shorthand);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-top", out var pt))
        {
            top = StyleParser.ParseLengthToTwips(pt);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-right", out var pr))
        {
            right = StyleParser.ParseLengthToTwips(pr);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-bottom", out var pb))
        {
            bottom = StyleParser.ParseLengthToTwips(pb);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-left", out var pl))
        {
            left = StyleParser.ParseLengthToTwips(pl);
            hasPadding = true;
        }

        if (!hasPadding)
        {
            return null;
        }

        return (top, right, bottom, left);
    }

    internal static T BuildMargin<T>(int? top, int? right, int? bottom, int? left)
        where T : OpenXmlCompositeElement, new()
    {
        var margin = new T();
        if (top != null)
        {
            margin.Append(new TopMargin
            {
                Width = top.Value.ToString(),
                Type = TableWidthUnitValues.Dxa
            });
        }

        if (left != null)
        {
            margin.Append(new StartMargin
            {
                Width = left.Value.ToString(),
                Type = TableWidthUnitValues.Dxa
            });
        }

        if (bottom != null)
        {
            margin.Append(new BottomMargin
            {
                Width = bottom.Value.ToString(),
                Type = TableWidthUnitValues.Dxa
            });
        }

        if (right != null)
        {
            margin.Append(new EndMargin
            {
                Width = right.Value.ToString(),
                Type = TableWidthUnitValues.Dxa
            });
        }

        return margin;
    }
}
