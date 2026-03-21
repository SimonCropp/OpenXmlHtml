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

    internal bool HasProperties =>
        MarginTopTwips != null ||
        MarginBottomTwips != null ||
        MarginLeftTwips != null ||
        MarginRightTwips != null ||
        TextIndentTwips != null ||
        LineHeightTwips != null ||
        LineHeightMultiple != null ||
        TextAlign != null ||
        BackgroundColor != null;
}
