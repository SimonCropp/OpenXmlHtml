class FormatState
{
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal bool Underline { get; set; }
    internal bool Strikethrough { get; set; }
    internal string? Color { get; set; }
    internal double? FontSizePt { get; set; }
    internal string? FontFamily { get; set; }
    internal bool Superscript { get; set; }
    internal bool Subscript { get; set; }
    internal ImageData? Image { get; set; }
    internal int ListDepth { get; set; }
    internal string? LinkUrl { get; set; }

    internal FormatState Copy() =>
        new()
        {
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Strikethrough = Strikethrough,
            Color = Color,
            FontSizePt = FontSizePt,
            FontFamily = FontFamily,
            Superscript = Superscript,
            Subscript = Subscript,
            Image = Image,
            ListDepth = ListDepth,
            LinkUrl = LinkUrl
        };

    internal bool HasFormatting =>
        Bold ||
        Italic ||
        Underline ||
        Strikethrough ||
        Superscript ||
        Subscript ||
        Color != null ||
        FontSizePt != null ||
        FontFamily != null;
}
