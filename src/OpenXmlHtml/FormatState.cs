class FormatState
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
    internal string? RunStyleId { get; set; }
    internal string? BackgroundColor { get; set; }
    internal BorderInfo? Border { get; set; }
    internal bool SmallCaps { get; set; }
    internal string? TextTransform { get; set; }

    internal FormatState Copy() =>
        new()
        {
            Bold = Bold,
            Italic = Italic,
            UnderlineStyle = UnderlineStyle,
            Strikethrough = Strikethrough,
            Color = Color,
            FontSizePt = FontSizePt,
            FontFamily = FontFamily,
            Superscript = Superscript,
            Subscript = Subscript,
            Image = Image,
            ListDepth = ListDepth,
            LinkUrl = LinkUrl,
            RunStyleId = RunStyleId,
            BackgroundColor = BackgroundColor,
            Border = Border,
            SmallCaps = SmallCaps,
            TextTransform = TextTransform
        };

    internal bool HasFormatting =>
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
