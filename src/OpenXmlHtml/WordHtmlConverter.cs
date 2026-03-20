using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlHtml;

/// <summary>
/// Converts HTML to OpenXml Wordprocessing elements for use in docx documents.
/// </summary>
public static class WordHtmlConverter
{
    /// <summary>
    /// Converts HTML to a list of Paragraph elements suitable for inserting into a Word document body.
    /// </summary>
    public static List<Paragraph> ToParagraphs(string html) =>
        ToParagraphs(html, null);

    /// <summary>
    /// Converts HTML to a list of Paragraph elements, embedding images into the given MainDocumentPart.
    /// </summary>
    public static List<Paragraph> ToParagraphs(string html, MainDocumentPart? mainPart)
    {
        var segments = HtmlSegmentParser.Parse(html);
        var paragraphs = new List<Paragraph>();
        var currentRuns = new List<OpenXmlElement>();
        var imageIndex = 0;
        var listDepth = 0;

        foreach (var segment in segments)
        {
            if (segment.Text == "\n")
            {
                paragraphs.Add(BuildParagraph(currentRuns, listDepth));
                currentRuns = [];
                listDepth = 0;
                continue;
            }

            if (segment.Format.Image != null)
            {
                if (mainPart != null)
                {
                    imageIndex++;
                    currentRuns.Add(BuildImageRun(mainPart, segment.Format.Image, imageIndex));
                }

                continue;
            }

            var text = segment.Text;
            if (segment.Format.ListDepth > 0)
            {
                listDepth = segment.Format.ListDepth;
                text = text.TrimStart();
            }

            var run = new Run();

            if (segment.Format.HasFormatting)
            {
                run.Append(BuildWordRunProperties(segment.Format));
            }

            run.Append(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            currentRuns.Add(run);
        }

        if (currentRuns.Count > 0)
        {
            paragraphs.Add(BuildParagraph(currentRuns, listDepth));
        }

        if (paragraphs.Count == 0)
        {
            paragraphs.Add(new());
        }

        return paragraphs;
    }

    /// <summary>
    /// Appends HTML content as paragraphs to a Word document body.
    /// </summary>
    public static void AppendHtml(Body body, string html) =>
        AppendHtml(body, html, null);

    /// <summary>
    /// Appends HTML content as paragraphs to a Word document body, embedding images into the given MainDocumentPart.
    /// </summary>
    public static void AppendHtml(Body body, string html, MainDocumentPart? mainPart)
    {
        foreach (var paragraph in ToParagraphs(html, mainPart))
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
        AppendHtml(body, html, mainPart);
        mainPart.Document = new(body);
    }

    /// <summary>
    /// Converts HTML from a stream to a docx file written to the given stream.
    /// </summary>
    public static void ConvertToDocx(Stream htmlStream, Stream docxStream)
    {
        using var reader = new StreamReader(htmlStream);
        var html = reader.ReadToEnd();
        ConvertToDocx(html, docxStream);
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

    static Paragraph BuildParagraph(List<OpenXmlElement> runs, int listDepth = 0)
    {
        var paragraph = new Paragraph();

        if (listDepth > 0)
        {
            paragraph.ParagraphProperties = new(
                new Indentation
                {
                    Left = (listDepth * 360).ToString()
                });
        }

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

    static Run BuildImageRun(MainDocumentPart mainPart, ImageData image, int imageIndex)
    {
        var imagePartType = GetImagePartType(image.ContentType);
        var relationshipId = $"rImage{imageIndex}";
        var imagePart = mainPart.AddImagePart(imagePartType, relationshipId);
        using (var ms = new MemoryStream(image.Bytes))
        {
            imagePart.FeedData(ms);
        }

        var widthEmu = (long)(image.WidthPx ?? 100) * 9525;
        var heightEmu = (long)(image.HeightPx ?? 100) * 9525;

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.DocProperties { Id = 1U, Name = "Image" },
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Image" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

        var run = new Run();
        run.Append(drawing);
        return run;
    }

    static string GetImagePartType(string contentType) =>
        contentType.ToLowerInvariant() switch
        {
            "image/jpeg" or "image/jpg" => "image/jpeg",
            "image/gif" => "image/gif",
            "image/bmp" => "image/bmp",
            "image/tiff" => "image/tiff",
            _ => "image/png"
        };
}
