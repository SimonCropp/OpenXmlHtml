namespace OpenXmlHtml;

class WordBuildContext
{
    internal List<OpenXmlElement> CurrentRuns = [];
    internal int ImageIndex;
    internal int ListDepth;
    internal int HeadingLevel;
    internal MainDocumentPart? MainPart;
}

static class WordContentBuilder
{
    static readonly HtmlParser parser = new();

    internal static List<OpenXmlElement> Build(string html, MainDocumentPart? mainPart)
    {
        var document = parser.ParseDocument($"<body>{html}</body>");
        var body = document.Body!;
        var elements = new List<OpenXmlElement>();
        var ctx = new WordBuildContext { MainPart = mainPart };
        ProcessChildren(body, new(), elements, ctx, false);
        FlushParagraph(elements, ctx);
        TrimTrailingEmptyParagraphs(elements);

        if (elements.Count == 0)
        {
            elements.Add(new Paragraph());
        }

        return elements;
    }

    static void ProcessChildren(INode node, FormatState format, List<OpenXmlElement> elements, WordBuildContext ctx, bool inPre)
    {
        foreach (var child in node.ChildNodes)
        {
            switch (child)
            {
                case IText textNode:
                {
                    var text = inPre ? textNode.Data : HtmlSegmentParser.CollapseWhitespace(textNode.Data);
                    if (text.Length > 0 &&
                        !(string.IsNullOrWhiteSpace(text) &&
                          HtmlSegmentParser.IsInterBlockWhitespace(textNode)))
                    {
                        AddTextRun(text, format, ctx);
                    }

                    break;
                }
                case IElement element:
                    ProcessElement(element, format, elements, ctx, inPre);
                    break;
            }
        }
    }

    static void ProcessElement(IElement element, FormatState format, List<OpenXmlElement> elements, WordBuildContext ctx, bool inPre)
    {
        var tag = element.LocalName;
        var newFormat = format.Copy();
        HtmlSegmentParser.ApplyElementFormatting(element, tag, newFormat);

        switch (tag)
        {
            case "br":
                ForceFlushParagraph(elements, ctx);
                return;
            case "hr":
                FlushParagraph(elements, ctx);
                AddTextRun("\u2014\u2014\u2014", format, ctx);
                FlushParagraph(elements, ctx);
                return;
            case "img":
            {
                var imageData = HtmlSegmentParser.ParseImageSrc(element);
                if (imageData != null)
                {
                    if (ctx.MainPart != null)
                    {
                        ctx.ImageIndex++;
                        ctx.CurrentRuns.Add(WordHtmlConverter.BuildImageRun(ctx.MainPart, imageData, ctx.ImageIndex));
                    }
                }
                else
                {
                    var alt = element.GetAttribute("alt");
                    if (!string.IsNullOrEmpty(alt))
                    {
                        // ReSharper disable once RedundantSuppressNullableWarningExpression
                        AddTextRun(alt!, format, ctx);
                    }
                }

                return;
            }
            case "svg":
            {
                if (ctx.MainPart != null)
                {
                    var imageData = HtmlSegmentParser.ParseSvgElement(element);
                    ctx.ImageIndex++;
                    ctx.CurrentRuns.Add(WordHtmlConverter.BuildImageRun(ctx.MainPart, imageData, ctx.ImageIndex));
                }

                return;
            }
            case "col":
                return;
            case "table":
                FlushParagraph(elements, ctx);
                BuildTable(element, format, elements, ctx);
                return;
            case "li":
            {
                FlushParagraph(elements, ctx);
                var parent = element.ParentElement?.LocalName;
                var depth = HtmlSegmentParser.GetListDepth(element);
                ctx.ListDepth = depth;

                if (parent == "ol")
                {
                    var index = 1;
                    foreach (var sibling in element.ParentElement!.Children)
                    {
                        if (sibling == element)
                        {
                            break;
                        }

                        if (sibling.LocalName == "li")
                        {
                            index++;
                        }
                    }

                    AddTextRun($"{index}. ", newFormat, ctx);
                }
                else
                {
                    var bullet = depth switch
                    {
                        0 => "\u25CF",
                        1 => "\u25CB",
                        _ => "\u25A0"
                    };
                    AddTextRun($"{bullet} ", newFormat, ctx);
                }

                ProcessChildren(element, newFormat, elements, ctx, inPre);
                FlushParagraph(elements, ctx);
                return;
            }
            case "a":
            {
                ProcessChildren(element, newFormat, elements, ctx, inPre);
                var href = element.GetAttribute("href");
                if (!string.IsNullOrEmpty(href))
                {
                    var linkText = element.TextContent.Trim();
                    if (linkText != href)
                    {
                        AddTextRun($" ({href})", format, ctx);
                    }
                }

                return;
            }
            case "q":
            {
                AddTextRun("\u201C", format, ctx);
                ProcessChildren(element, newFormat, elements, ctx, inPre);
                AddTextRun("\u201D", format, ctx);
                return;
            }
            case "pre":
            {
                FlushParagraph(elements, ctx);
                ProcessChildren(element, newFormat, elements, ctx, true);
                FlushParagraph(elements, ctx);
                return;
            }
        }

        var isBlock = HtmlSegmentParser.IsBlockElement(tag);
        var pageBreakBefore = false;
        var pageBreakAfter = false;

        if (isBlock)
        {
            var style = element.GetAttribute("style");
            if (style != null)
            {
                var declarations = StyleParser.Parse(style);
                pageBreakBefore = declarations.TryGetValue("page-break-before", out var pbb) && pbb == "always";
                pageBreakAfter = declarations.TryGetValue("page-break-after", out var pba) && pba == "always";
            }

            FlushParagraph(elements, ctx);

            if (tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
            {
                ctx.HeadingLevel = tag[1] - '0';
            }

            if (pageBreakBefore)
            {
                elements.Add(new Paragraph(
                    new ParagraphProperties(new PageBreakBefore())));
            }
        }

        ProcessChildren(element, newFormat, elements, ctx, inPre);

        if (isBlock)
        {
            FlushParagraph(elements, ctx);

            if (pageBreakAfter)
            {
                elements.Add(new Paragraph(
                    new ParagraphProperties(new PageBreakBefore())));
            }
        }
    }

    static void AddTextRun(string text, FormatState format, WordBuildContext ctx)
    {
        var run = new Run();

        if (format.HasFormatting)
        {
            run.Append(WordHtmlConverter.BuildWordRunProperties(format));
        }

        run.Append(
            new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        ctx.CurrentRuns.Add(run);
    }

    static void FlushParagraph(List<OpenXmlElement> elements, WordBuildContext ctx)
    {
        if (ctx.CurrentRuns.Count == 0)
        {
            ctx.HeadingLevel = 0;
            return;
        }

        var paragraph = WordHtmlConverter.BuildParagraph(ctx.CurrentRuns, ctx.ListDepth);
        if (ctx.HeadingLevel > 0)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId = new() { Val = $"Heading{ctx.HeadingLevel}" };
        }

        elements.Add(paragraph);
        ctx.CurrentRuns = [];
        ctx.ListDepth = 0;
        ctx.HeadingLevel = 0;
    }

    static void ForceFlushParagraph(List<OpenXmlElement> elements, WordBuildContext ctx)
    {
        elements.Add(WordHtmlConverter.BuildParagraph(ctx.CurrentRuns, ctx.ListDepth));
        ctx.CurrentRuns = [];
        ctx.ListDepth = 0;
    }

    static void TrimTrailingEmptyParagraphs(List<OpenXmlElement> elements)
    {
        while (elements.Count > 0 && elements[^1] is Paragraph { HasChildren: false })
        {
            elements.RemoveAt(elements.Count - 1);
        }
    }

    static void BuildTable(IElement tableElement, FormatState format, List<OpenXmlElement> elements, WordBuildContext ctx)
    {
        // Handle caption before the table
        IElement? caption = null;
        foreach (var child in tableElement.Children)
        {
            if (child.LocalName == "caption")
            {
                caption = child;
                break;
            }
        }

        if (caption != null)
        {
            var captionFormat = format.Copy();
            HtmlSegmentParser.ApplyElementFormatting(caption, "caption", captionFormat);
            ProcessChildren(caption, captionFormat, elements, ctx, false);
            FlushParagraph(elements, ctx);
        }

        var rows = GetTableRows(tableElement);
        if (rows.Count == 0)
        {
            return;
        }

        var columnCount = GetColumnCount(rows);
        var table = new Table();

        table.Append(
            new TableProperties(
                new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" })));

        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.Append(new GridColumn());
        }

        table.Append(grid);

        // Track rowspans: starting column index -> (remaining rows, colspan)
        var rowspanTracker = new Dictionary<int, (int Remaining, int Colspan)>();

        foreach (var row in rows)
        {
            var tr = new TableRow();
            var cells = GetCells(row);
            var cellIndex = 0;
            var colIndex = 0;

            while (colIndex < columnCount)
            {
                if (rowspanTracker.TryGetValue(colIndex, out var spanInfo))
                {
                    var contTcPr = new TableCellProperties(new VerticalMerge());
                    if (spanInfo.Colspan > 1)
                    {
                        contTcPr.Append(new GridSpan { Val = spanInfo.Colspan });
                    }

                    tr.Append(new TableCell(contTcPr, new Paragraph()));

                    if (spanInfo.Remaining <= 1)
                    {
                        rowspanTracker.Remove(colIndex);
                    }
                    else
                    {
                        rowspanTracker[colIndex] = (spanInfo.Remaining - 1, spanInfo.Colspan);
                    }

                    colIndex += spanInfo.Colspan;
                    continue;
                }

                if (cellIndex >= cells.Count)
                {
                    tr.Append(new TableCell(new Paragraph()));
                    colIndex++;
                    continue;
                }

                var cellElement = cells[cellIndex];
                cellIndex++;

                var colspan = ParseIntAttribute(cellElement, "colspan", 1);
                var rowspan = ParseIntAttribute(cellElement, "rowspan", 1);

                tr.Append(BuildTableCell(cellElement, format, ctx, colspan, rowspan > 1));

                if (rowspan > 1)
                {
                    rowspanTracker[colIndex] = (rowspan - 1, colspan);
                }

                colIndex += colspan;
            }

            table.Append(tr);
        }

        elements.Add(table);
    }

    static TableCell BuildTableCell(IElement cellElement, FormatState format, WordBuildContext parentCtx, int colspan, bool isRowspanStart)
    {
        var tc = new TableCell();

        if (colspan > 1 || isRowspanStart)
        {
            var tcPr = new TableCellProperties();
            if (colspan > 1)
            {
                tcPr.Append(new GridSpan { Val = colspan });
            }

            if (isRowspanStart)
            {
                tcPr.Append(new VerticalMerge { Val = MergedCellValues.Restart });
            }

            tc.Append(tcPr);
        }

        var cellFormat = format.Copy();
        HtmlSegmentParser.ApplyElementFormatting(cellElement, cellElement.LocalName, cellFormat);

        var cellElements = new List<OpenXmlElement>();
        var cellCtx = new WordBuildContext { MainPart = parentCtx.MainPart, ImageIndex = parentCtx.ImageIndex };
        ProcessChildren(cellElement, cellFormat, cellElements, cellCtx, false);
        FlushParagraph(cellElements, cellCtx);
        parentCtx.ImageIndex = cellCtx.ImageIndex;

        if (cellElements.Count == 0)
        {
            tc.Append(new Paragraph());
        }
        else
        {
            foreach (var el in cellElements)
            {
                tc.Append(el);
            }
        }

        return tc;
    }

    static List<IElement> GetTableRows(IElement tableElement)
    {
        var rows = new List<IElement>();
        foreach (var child in tableElement.Children)
        {
            switch (child.LocalName)
            {
                case "tr":
                    rows.Add(child);
                    break;
                case "thead" or "tbody" or "tfoot":
                {
                    foreach (var grandchild in child.Children)
                    {
                        if (grandchild.LocalName == "tr")
                        {
                            rows.Add(grandchild);
                        }
                    }

                    break;
                }
            }
        }

        return rows;
    }

    static List<IElement> GetCells(IElement row)
    {
        var cells = new List<IElement>();
        foreach (var child in row.Children)
        {
            if (child.LocalName is "td" or "th")
            {
                cells.Add(child);
            }
        }

        return cells;
    }

    static int GetColumnCount(List<IElement> rows)
    {
        var maxCols = 0;
        foreach (var row in rows)
        {
            var cols = 0;
            foreach (var cell in row.Children)
            {
                if (cell.LocalName is "td" or "th")
                {
                    cols += ParseIntAttribute(cell, "colspan", 1);
                }
            }

            maxCols = Math.Max(maxCols, cols);
        }

        return Math.Max(1, maxCols);
    }

    static int ParseIntAttribute(IElement element, string attribute, int defaultValue)
    {
        var value = element.GetAttribute(attribute);
        if (value != null && int.TryParse(value, out var result) && result > 1)
        {
            return result;
        }

        return defaultValue;
    }
}
