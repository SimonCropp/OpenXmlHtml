using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlHtml;

/// <summary>
/// Converts HTML to OpenXml Spreadsheet InlineString elements for use in xlsx cells.
/// </summary>
public static class SpreadsheetHtmlConverter
{
    /// <summary>
    /// Converts an HTML string to an InlineString containing formatted Run elements.
    /// </summary>
    public static InlineString ToInlineString(string html) =>
        ToInlineString(HtmlSegmentParser.Parse(html));

    /// <summary>
    /// Sets the value of a spreadsheet cell from HTML content.
    /// </summary>
    public static void SetCellHtml(SpreadsheetCell cell, string html)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = ToInlineString(html);
    }

    /// <summary>
    /// Sets the value of a spreadsheet cell from HTML content, applying wrap text and hyperlinks.
    /// </summary>
    public static void SetCellHtml(SpreadsheetCell cell, string html, WorksheetPart worksheetPart)
    {
        var segments = HtmlSegmentParser.Parse(html);
        cell.DataType = CellValues.InlineString;
        cell.InlineString = ToInlineString(segments);

        var workbookPart = ((SpreadsheetDocument) worksheetPart.OpenXmlPackage).WorkbookPart!;

        if (HasNewlines(cell.InlineString))
        {
            cell.StyleIndex = EnsureWrapTextStyle(workbookPart);
        }

        var linkUrl = GetSingleLinkUrl(segments);
        if (linkUrl != null && cell.CellReference?.Value != null)
        {
            ApplyHyperlink(worksheetPart, cell.CellReference!, linkUrl);
        }
    }

    /// <summary>
    /// Converts an HTML string to an InlineString, with settings controlling remote image resolution.
    /// </summary>
    public static InlineString ToInlineString(string html, HtmlConvertSettings settings) =>
        ToInlineString(HtmlSegmentParser.Parse(html, settings));

    /// <summary>
    /// Sets the value of a spreadsheet cell from HTML content, with settings controlling remote image resolution.
    /// </summary>
    public static void SetCellHtml(SpreadsheetCell cell, string html, HtmlConvertSettings settings)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = ToInlineString(html, settings);
    }

    /// <summary>
    /// Sets the value of a spreadsheet cell from HTML content, with settings controlling remote image resolution, applying wrap text and hyperlinks.
    /// </summary>
    public static void SetCellHtml(SpreadsheetCell cell, string html, WorksheetPart worksheetPart, HtmlConvertSettings settings)
    {
        var segments = HtmlSegmentParser.Parse(html, settings);
        cell.DataType = CellValues.InlineString;
        cell.InlineString = ToInlineString(segments);

        var workbookPart = ((SpreadsheetDocument) worksheetPart.OpenXmlPackage).WorkbookPart!;

        if (HasNewlines(cell.InlineString))
        {
            cell.StyleIndex = EnsureWrapTextStyle(workbookPart);
        }

        var linkUrl = GetSingleLinkUrl(segments);
        if (linkUrl != null && cell.CellReference?.Value != null)
        {
            ApplyHyperlink(worksheetPart, cell.CellReference!, linkUrl);
        }
    }

    static InlineString ToInlineString(List<TextSegment> segments)
    {
        var inlineString = new InlineString();

        foreach (var segment in segments)
        {
            if (segment.Format.Image != null)
            {
                continue;
            }

            var run = new SpreadsheetRun();

            if (segment.Format.HasFormatting)
            {
                run.Append(BuildSpreadsheetRunProperties(segment.Format));
            }

            run.Append(
                new SpreadsheetText(XmlCharFilter.StripInvalidXmlChars(segment.Text))
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            inlineString.Append(run);
        }

        return inlineString;
    }

    static string? GetSingleLinkUrl(List<TextSegment> segments)
    {
        string? url = null;
        foreach (var segment in segments)
        {
            if (segment.Format.LinkUrl == null)
            {
                continue;
            }

            if (url != null &&
                url != segment.Format.LinkUrl)
            {
                return null;
            }

            url = segment.Format.LinkUrl;
        }

        return url;
    }

    static void ApplyHyperlink(WorksheetPart worksheetPart, string cellReference, string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return;
        }

        var relationship = worksheetPart.AddHyperlinkRelationship(uri, true);

        var worksheet = worksheetPart.Worksheet!;
        var hyperlinks = worksheet.GetFirstChild<Hyperlinks>();
        if (hyperlinks == null)
        {
            hyperlinks = new();
            var sheetData = worksheet.GetFirstChild<SheetData>()!;
            worksheet.InsertAfter(hyperlinks, sheetData);
        }

        hyperlinks.Append(
            new DocumentFormat.OpenXml.Spreadsheet.Hyperlink
            {
                Reference = cellReference,
                Id = relationship.Id
            });
    }

    static bool HasNewlines(InlineString inlineString)
    {
        foreach (var text in inlineString.Descendants<SpreadsheetText>())
        {
            if (text.Text.Contains('\n'))
            {
                return true;
            }
        }

        return false;
    }

    static uint EnsureWrapTextStyle(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.WorkbookStylesPart
                         ?? workbookPart.AddNewPart<WorkbookStylesPart>();

        var stylesheet = stylesPart.Stylesheet;
        if (stylesheet == null)
        {
            stylesheet = new();
            stylesPart.Stylesheet = stylesheet;
        }

        stylesheet.Fonts ??=
            new(new DocumentFormat.OpenXml.Spreadsheet.Font())
            {
                Count = 1
            };
        stylesheet.Fills ??=
            new(
                new Fill(
                    new PatternFill
                    {
                        PatternType = PatternValues.None
                    }),
                new Fill(
                    new PatternFill
                    {
                        PatternType = PatternValues.Gray125
                    })
            )
            {
                Count = 2
            };
        stylesheet.Borders ??=
            new(
                new DocumentFormat.OpenXml.Spreadsheet.Border())
            {
                Count = 1
            };
        stylesheet.CellFormats ??=
            new(new CellFormat())
            {
                Count = 1
            };

        uint index = 0;
        foreach (var cf in stylesheet.CellFormats.Elements<CellFormat>())
        {
            if (cf.GetFirstChild<Alignment>()?.WrapText?.Value == true)
            {
                return index;
            }

            index++;
        }

        stylesheet.CellFormats.Append(
            new CellFormat(
                new Alignment
                {
                    WrapText = true
                })
            {
                ApplyAlignment = true
            });
        stylesheet.CellFormats.Count = index + 1;
        return index;
    }

    static SpreadsheetRunProperties BuildSpreadsheetRunProperties(FormatState format)
    {
        var props = new SpreadsheetRunProperties();

        if (format.Bold)
        {
            props.Append(new SpreadsheetBold());
        }

        if (format.Italic)
        {
            props.Append(new SpreadsheetItalic());
        }

        if (format.UnderlineStyle != null)
        {
            props.Append(new SpreadsheetUnderline());
        }

        if (format.Strikethrough)
        {
            props.Append(new SpreadsheetStrike());
        }

        if (format.Color != null)
        {
            props.Append(
                new SpreadsheetColor
                {
                    Rgb = string.Concat("FF", format.Color)
                });
        }

        if (format.FontSizePt != null)
        {
            props.Append(
                new SpreadsheetFontSize
                {
                    Val = format.FontSizePt.Value
                });
        }

        if (format.FontFamily != null)
        {
            props.Append(
                new SpreadsheetRunFont
                {
                    Val = format.FontFamily
                });
        }

        if (format.Superscript)
        {
            props.Append(
                new DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment
                {
                    Val = VerticalAlignmentRunValues.Superscript
                });
        }
        else if (format.Subscript)
        {
            props.Append(
                new DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment
                {
                    Val = VerticalAlignmentRunValues.Subscript
                });
        }

        return props;
    }
}
