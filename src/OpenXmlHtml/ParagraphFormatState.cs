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
}
