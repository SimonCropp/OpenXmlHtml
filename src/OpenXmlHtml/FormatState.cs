struct FormatState
{
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal UnderlineValues? UnderlineStyle { get; set; }
    internal bool Strikethrough { get; set; }
    internal string? Color { get; set; }
    internal double? FontSizePt { get; set; }
    internal string? FontFamily { get; set; }
    internal bool Superscript { get; set; }
    internal bool Subscript { get; set; }
    internal ImageData? Image { get; set; }
    internal int ListDepth { get; set; }
    internal string? LinkUrl { get; set; }
    internal string? LinkTitle { get; set; }
    internal string? RunStyleId { get; set; }
    internal string? BackgroundColor { get; set; }
    internal BorderInfo? Border { get; set; }
    internal bool SmallCaps { get; set; }
    internal string? TextTransform { get; set; }

    internal readonly bool HasFormatting =>
        Bold ||
        Italic ||
        UnderlineStyle != null ||
        Strikethrough ||
        Superscript ||
        Subscript ||
        Color != null ||
        FontSizePt != null ||
        FontFamily != null ||
        RunStyleId != null ||
        BackgroundColor != null ||
        Border != null ||
        SmallCaps;
}
