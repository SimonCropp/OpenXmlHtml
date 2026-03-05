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
