namespace OpenXmlHtml;

/// <summary>
/// Converts HTML to OpenXml Wordprocessing elements for use in docx documents.
/// </summary>
public static class WordHtmlConverter
{
    /// <summary>
    /// Converts HTML to a list of Paragraph elements suitable for inserting into a Word document body.
    /// </summary>
    public static List<Paragraph> ToParagraphs(string html)
    {
        var segments = HtmlSegmentParser.Parse(html);
        var paragraphs = new List<Paragraph>();
        var currentRuns = new List<Run>();

        foreach (var segment in segments)
        {
            if (segment.Text == "\n")
            {
                paragraphs.Add(BuildParagraph(currentRuns));
                currentRuns = [];
                continue;
            }

            var run = new Run();

            if (segment.Format.HasFormatting)
            {
                run.Append(BuildWordRunProperties(segment.Format));
            }

            run.Append(
                new Text(segment.Text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            currentRuns.Add(run);
        }

        if (currentRuns.Count > 0)
        {
            paragraphs.Add(BuildParagraph(currentRuns));
        }

        if (paragraphs.Count == 0)
        {
            paragraphs.Add(new Paragraph());
        }

        return paragraphs;
    }

    /// <summary>
    /// Appends HTML content as paragraphs to a Word document body.
    /// </summary>
    public static void AppendHtml(Body body, string html)
    {
        foreach (var paragraph in ToParagraphs(html))
        {
            body.Append(paragraph);
        }
    }

    /// <summary>
    /// Converts an HTML string to a docx file written to the given stream.
    /// </summary>
    public static void ConvertToDocx(string html, Stream stream)
    {
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();
        AppendHtml(body, html);
        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(body);
    }

    /// <summary>
    /// Converts an HTML file to a docx file.
    /// </summary>
    public static void ConvertFileToDocx(string htmlPath, string docxPath)
    {
        var html = File.ReadAllText(htmlPath);
        using var stream = File.Create(docxPath);
        ConvertToDocx(html, stream);
    }

    static Paragraph BuildParagraph(List<Run> runs)
    {
        var paragraph = new Paragraph();
        foreach (var run in runs)
        {
            paragraph.Append(run);
        }

        return paragraph;
    }

    static RunProperties BuildWordRunProperties(FormatState format)
    {
        var props = new RunProperties();

        if (format.Bold)
        {
            props.Append(new Bold());
        }

        if (format.Italic)
        {
            props.Append(new Italic());
        }

        if (format.Underline)
        {
            props.Append(new Underline { Val = UnderlineValues.Single });
        }

        if (format.Strikethrough)
        {
            props.Append(new Strike());
        }

        if (format.Color != null)
        {
            props.Append(new Color { Val = format.Color });
        }

        if (format.FontSizePt != null)
        {
            var halfPoints = (int)(format.FontSizePt.Value * 2);
            props.Append(new FontSize { Val = halfPoints.ToString() });
        }

        if (format.FontFamily != null)
        {
            props.Append(new RunFonts { Ascii = format.FontFamily, HighAnsi = format.FontFamily });
        }

        if (format.Superscript)
        {
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        }
        else if (format.Subscript)
        {
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        }

        return props;
    }
}
