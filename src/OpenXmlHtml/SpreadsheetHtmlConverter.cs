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
    public static InlineString ToInlineString(string html)
    {
        var segments = HtmlSegmentParser.Parse(html);
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
                new SpreadsheetText(segment.Text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            inlineString.Append(run);
        }

        return inlineString;
    }

    /// <summary>
    /// Sets the value of a spreadsheet cell from HTML content.
    /// </summary>
    public static void SetCellHtml(SpreadsheetCell cell, string html)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = ToInlineString(html);
    }

    /// <summary>
    /// Sets the value of a spreadsheet cell from HTML content, applying wrap text when the content contains newlines.
    /// </summary>
    public static void SetCellHtml(SpreadsheetCell cell, string html, WorkbookPart workbookPart)
    {
        SetCellHtml(cell, html);

        if (HasNewlines(cell.InlineString!))
        {
            cell.StyleIndex = EnsureWrapTextStyle(workbookPart);
        }
    }

    static bool HasNewlines(InlineString inlineString)
    {
        foreach (var text in inlineString.Descendants<SpreadsheetText>())
        {
            if (text.Text?.Contains('\n') == true)
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
            stylesheet = new Stylesheet();
            stylesPart.Stylesheet = stylesheet;
        }

        stylesheet.Fonts ??= new DocumentFormat.OpenXml.Spreadsheet.Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1 };
        stylesheet.Fills ??= new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
        )
        { Count = 2 };
        stylesheet.Borders ??= new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()) { Count = 1 };
        stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1 };

        uint index = 0;
        foreach (var cf in stylesheet.CellFormats.Elements<CellFormat>())
        {
            if (cf.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Alignment>()?.WrapText?.Value == true)
            {
                return index;
            }

            index++;
        }

        stylesheet.CellFormats.Append(
            new CellFormat(new DocumentFormat.OpenXml.Spreadsheet.Alignment { WrapText = true })
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

        if (format.Underline)
        {
            props.Append(new SpreadsheetUnderline());
        }

        if (format.Strikethrough)
        {
            props.Append(new SpreadsheetStrike());
        }

        if (format.Color != null)
        {
            props.Append(new SpreadsheetColor { Rgb = "FF" + format.Color });
        }

        if (format.FontSizePt != null)
        {
            props.Append(new SpreadsheetFontSize { Val = format.FontSizePt.Value });
        }

        if (format.FontFamily != null)
        {
            props.Append(new SpreadsheetRunFont { Val = format.FontFamily });
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
