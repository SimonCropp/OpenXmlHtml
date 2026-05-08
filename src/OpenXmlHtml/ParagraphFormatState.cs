class ParagraphFormatState
{
    internal int? MarginTopTwips { get; set; }
    internal int? MarginBottomTwips { get; set; }
    internal int? MarginLeftTwips { get; set; }
    internal int? MarginRightTwips { get; set; }
    internal int? TextIndentTwips { get; set; }
    internal int? LineHeightTwips { get; set; }
    internal double? LineHeightMultiple { get; set; }
    internal JustificationValues? TextAlign { get; set; }
    internal string? BackgroundColor { get; set; }
    internal TextDirectionValues? WritingMode { get; set; }
    internal BorderInfo? BorderTop { get; set; }
    internal BorderInfo? BorderRight { get; set; }
    internal BorderInfo? BorderBottom { get; set; }
    internal BorderInfo? BorderLeft { get; set; }

    internal bool HasProperties =>
        MarginTopTwips != null ||
        MarginBottomTwips != null ||
        MarginLeftTwips != null ||
        MarginRightTwips != null ||
        TextIndentTwips != null ||
        LineHeightTwips != null ||
        LineHeightMultiple != null ||
        TextAlign != null ||
        BackgroundColor != null ||
        WritingMode != null ||
        BorderTop != null ||
        BorderRight != null ||
        BorderBottom != null ||
        BorderLeft != null;

    internal static ParagraphFormatState ParseFrom(Dictionary<string, string> declarations)
    {
        var pf = new ParagraphFormatState();

        if (declarations.TryGetValue("margin", out var marginShorthand))
        {
            var (t, r, b, l) = StyleParser.ParseMarginShorthand(marginShorthand);
            pf.MarginTopTwips = t;
            pf.MarginRightTwips = r;
            pf.MarginBottomTwips = b;
            pf.MarginLeftTwips = l;
        }

        if (declarations.TryGetValue("margin-top", out var mt))
        {
            pf.MarginTopTwips = StyleParser.ParseLengthToTwips(mt);
        }

        if (declarations.TryGetValue("margin-bottom", out var mb))
        {
            pf.MarginBottomTwips = StyleParser.ParseLengthToTwips(mb);
        }

        if (declarations.TryGetValue("margin-left", out var ml))
        {
            pf.MarginLeftTwips = StyleParser.ParseLengthToTwips(ml);
        }

        if (declarations.TryGetValue("margin-right", out var mr))
        {
            pf.MarginRightTwips = StyleParser.ParseLengthToTwips(mr);
        }

        if (declarations.TryGetValue("text-indent", out var ti))
        {
            pf.TextIndentTwips = StyleParser.ParseLengthToTwips(ti);
        }

        if (declarations.TryGetValue("text-align", out var ta))
        {
            pf.TextAlign = StyleParser.ParseTextAlign(ta);
        }

        if (declarations.TryGetValue("background-color", out var blockBg) ||
            declarations.TryGetValue("background", out blockBg))
        {
            var parsed = ColorParser.Parse(blockBg);
            if (parsed != null)
            {
                pf.BackgroundColor = parsed;
            }
        }

        if (declarations.TryGetValue("line-height", out var lh))
        {
            ParseLineHeight(lh, pf);
        }

        if (declarations.TryGetValue("writing-mode", out var writingMode))
        {
            pf.WritingMode = writingMode.Equals("vertical-rl", StringComparison.OrdinalIgnoreCase) || writingMode.Equals("tb-rl", StringComparison.OrdinalIgnoreCase)
                ? TextDirectionValues.TopToBottomRightToLeft
                : writingMode.Equals("vertical-lr", StringComparison.OrdinalIgnoreCase) || writingMode.Equals("tb-lr", StringComparison.OrdinalIgnoreCase)
                    ? TextDirectionValues.BottomToTopLeftToRight
                    : null;
        }

        if (declarations.TryGetValue("direction", out var direction) &&
            direction.Equals("rtl", StringComparison.OrdinalIgnoreCase))
        {
            pf.WritingMode ??= TextDirectionValues.TopToBottomRightToLeft;
        }

        if (declarations.TryGetValue("border", out var borderAll))
        {
            var bi = StyleParser.ParseBorder(borderAll);
            pf.BorderTop = bi;
            pf.BorderRight = bi;
            pf.BorderBottom = bi;
            pf.BorderLeft = bi;
        }

        if (declarations.TryGetValue("border-top", out var bt))
        {
            pf.BorderTop = StyleParser.ParseBorder(bt);
        }

        if (declarations.TryGetValue("border-right", out var br))
        {
            pf.BorderRight = StyleParser.ParseBorder(br);
        }

        if (declarations.TryGetValue("border-bottom", out var bb))
        {
            pf.BorderBottom = StyleParser.ParseBorder(bb);
        }

        if (declarations.TryGetValue("border-left", out var bl))
        {
            pf.BorderLeft = StyleParser.ParseBorder(bl);
        }

        return pf;
    }

    static void ParseLineHeight(string lh, ParagraphFormatState pf)
    {
        if (lh.EndsWith('%'))
        {
            if (double.TryParse(lh.AsSpan()[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct))
            {
                pf.LineHeightMultiple = pct / 100.0;
            }
        }
        else if (double.TryParse(lh, NumberStyles.Float, CultureInfo.InvariantCulture, out var multiple))
        {
            // NumberStyles.Float rejects any non-numeric trailing chars, so a
            // successful parse already excludes "pt", "px", "em" suffixes.
            pf.LineHeightMultiple = multiple;
        }
        else
        {
            pf.LineHeightTwips = StyleParser.ParseLengthToTwips(lh);
        }
    }
}
