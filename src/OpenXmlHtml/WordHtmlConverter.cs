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
    /// Converts HTML to a list of OpenXml elements (paragraphs, tables, etc.) for inserting into a Word document body.
    /// </summary>
    public static List<OpenXmlElement> ToElements(string html) =>
        ToElements(html, null);

    /// <summary>
    /// Converts HTML to a list of OpenXml elements, embedding images into the given MainDocumentPart.
    /// </summary>
    public static List<OpenXmlElement> ToElements(string html, MainDocumentPart? mainPart) =>
        WordContentBuilder.Build(html, mainPart);

    /// <summary>
    /// Appends HTML content as paragraphs to a Word document body.
    /// </summary>
    public static void AppendHtml(Body body, string html) =>
        AppendHtml(body, html, null);

    /// <summary>
    /// Appends HTML content to a Word document body, embedding images into the given MainDocumentPart.
    /// </summary>
    public static void AppendHtml(Body body, string html, MainDocumentPart? mainPart)
    {
        foreach (var element in ToElements(html, mainPart))
        {
            body.Append(element);
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

    internal static Paragraph BuildParagraph(List<OpenXmlElement> runs, int listDepth = 0)
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

    internal static RunProperties BuildWordRunProperties(FormatState format)
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

    // Small PNG used as fallback for SVG images in older Word versions
    static readonly byte[] fallbackPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==");

    internal static Run BuildImageRun(MainDocumentPart mainPart, ImageData image, int imageIndex)
    {
        if (image.ContentType == "image/svg+xml")
        {
            return BuildSvgImageRun(mainPart, image, imageIndex);
        }

        var imagePartType = GetImagePartType(image.ContentType);
        var relationshipId = $"rImage{imageIndex}";
        var imagePart = mainPart.AddImagePart(imagePartType, relationshipId);
        using (var ms = new MemoryStream(image.Bytes))
        {
            imagePart.FeedData(ms);
        }

        var widthEmu = (long)(image.WidthPx ?? 100) * 9525;
        var heightEmu = (long)(image.HeightPx ?? 100) * 9525;

        var drawing = BuildImageDrawing(
            new()
            {
                Embed = relationshipId
            },
            widthEmu,
            heightEmu);

        var run = new Run();
        run.Append(drawing);
        return run;
    }

    static Run BuildSvgImageRun(MainDocumentPart mainPart, ImageData image, int imageIndex)
    {
        // Add PNG fallback for older Word versions
        var pngRelId = $"rImage{imageIndex}";
        var pngPart = mainPart.AddImagePart("image/png", pngRelId);
        using (var ms = new MemoryStream(fallbackPng))
        {
            pngPart.FeedData(ms);
        }

        // Add SVG image part
        var svgRelId = $"rImage{imageIndex}s";
        var svgPart = mainPart.AddImagePart("image/svg+xml", svgRelId);
        using (var ms = new MemoryStream(image.Bytes))
        {
            svgPart.FeedData(ms);
        }

        var widthEmu = (long)(image.WidthPx ?? 100) * 9525;
        var heightEmu = (long)(image.HeightPx ?? 100) * 9525;

        // Blip references PNG fallback, with SVG extension for Word 2016+
        var blip = new A.Blip { Embed = pngRelId };
        var svgBlipElement = new OpenXmlUnknownElement("asvg", "svgBlip", "http://schemas.microsoft.com/office/drawing/2016/SVG/main");
        svgBlipElement.SetAttribute(new("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", svgRelId));
        var ext = new OpenXmlUnknownElement("a", "ext", "http://schemas.openxmlformats.org/drawingml/2006/main");
        ext.SetAttribute(new("", "uri", "", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"));
        ext.Append(svgBlipElement);
        var extList = new OpenXmlUnknownElement("a", "extLst", "http://schemas.openxmlformats.org/drawingml/2006/main");
        extList.Append(ext);
        blip.Append(extList);

        var drawing = BuildImageDrawing(blip, widthEmu, heightEmu);

        var run = new Run();
        run.Append(drawing);
        return run;
    }

    static Drawing BuildImageDrawing(A.Blip blip, long widthEmu, long heightEmu) =>
        new(
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
                                blip,
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
