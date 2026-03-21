namespace OpenXmlHtml;

class WordBuildContext
{
    internal List<OpenXmlElement> CurrentRuns = [];
    internal int ImageIndex;
    internal int ListDepth;
    internal int HeadingLevel;
    internal int FootnoteIndex;
    internal int BookmarkId;
    internal MainDocumentPart? MainPart;
    internal HtmlConvertSettings? Settings;
    internal Dictionary<string, StyleType>? StyleMap;
    internal string? ParagraphStyleId;
    internal ParagraphFormatState? ParagraphFormat;
    internal Stack<(int NumId, int Ilvl, bool IsOrdered)> ListStack = new();
    internal int? BulletAbstractNumId;
    internal int? DecimalAbstractNumId;
    internal int NextNumId;
    internal int? ListNumId;
    internal int? ListIlvl;
}

static class WordContentBuilder
{
    static readonly HtmlParser parser = new();

    internal static List<OpenXmlElement> Build(string html, MainDocumentPart? mainPart, HtmlConvertSettings? settings = null)
    {
        var document = parser.ParseDocument($"<body>{html}</body>");
        var body = document.Body!;
        var elements = new List<OpenXmlElement>();
        var ctx = new WordBuildContext
        {
            MainPart = mainPart,
            Settings = settings,
            StyleMap = WordStyleLookup.BuildStyleMap(mainPart)
        };
        if (mainPart?.NumberingDefinitionsPart?.Numbering is { } existingNumbering)
        {
            ctx.NextNumId = WordNumberingBuilder.GetNextId(existingNumbering);
        }
        else
        {
            ctx.NextNumId = 1;
        }

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
                var imageData = ImageResolver.Resolve(element, ctx.Settings);
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
            case "ul" or "ol":
            {
                FlushParagraph(elements, ctx);
                if (ctx.MainPart != null)
                {
                    var isOrdered = tag == "ol";
                    var part = WordNumberingBuilder.EnsureNumberingPart(ctx.MainPart);
                    var numbering = part.Numbering!;
                    int abstractNumId;
                    if (isOrdered)
                    {
                        if (ctx.DecimalAbstractNumId == null)
                        {
                            var id = WordNumberingBuilder.GetNextId(numbering);
                            ctx.DecimalAbstractNumId = WordNumberingBuilder.CreateDecimalAbstractNum(numbering, id);
                        }

                        abstractNumId = ctx.DecimalAbstractNumId.Value;
                    }
                    else
                    {
                        if (ctx.BulletAbstractNumId == null)
                        {
                            var id = WordNumberingBuilder.GetNextId(numbering);
                            ctx.BulletAbstractNumId = WordNumberingBuilder.CreateBulletAbstractNum(numbering, id);
                        }

                        abstractNumId = ctx.BulletAbstractNumId.Value;
                    }

                    var numId = WordNumberingBuilder.GetNextId(numbering);
                    WordNumberingBuilder.AddNumberingInstance(numbering, numId, abstractNumId);
                    var ilvl = ctx.ListStack.Count > 0 ? ctx.ListStack.Peek().Ilvl + 1 : 0;
                    ctx.ListStack.Push((numId, ilvl, isOrdered));
                    ProcessChildren(element, newFormat, elements, ctx, inPre);
                    ctx.ListStack.Pop();
                }
                else
                {
                    ProcessChildren(element, newFormat, elements, ctx, inPre);
                }

                FlushParagraph(elements, ctx);
                return;
            }
            case "li":
            {
                FlushParagraph(elements, ctx);
                if (ctx.ListStack.Count > 0)
                {
                    var (numId, ilvl, _) = ctx.ListStack.Peek();
                    ctx.ListNumId = numId;
                    ctx.ListIlvl = ilvl;
                }
                else
                {
                    // Fallback: text prefix when no MainDocumentPart
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
                }

                ProcessChildren(element, newFormat, elements, ctx, inPre);
                FlushParagraph(elements, ctx);
                return;
            }
            case "a":
            {
                var href = element.GetAttribute("href");
                if (href != null && href.StartsWith('#') && href.Length > 1)
                {
                    // Internal anchor link → Word hyperlink to bookmark
                    var runsBefore = ctx.CurrentRuns.Count;
                    ProcessChildren(element, newFormat, elements, ctx, inPre);
                    var hyperlink = new Hyperlink { Anchor = href[1..] };
                    // Move newly added runs into the hyperlink
                    while (ctx.CurrentRuns.Count > runsBefore)
                    {
                        var run = ctx.CurrentRuns[runsBefore];
                        ctx.CurrentRuns.RemoveAt(runsBefore);
                        hyperlink.Append(run);
                    }

                    ctx.CurrentRuns.Add(hyperlink);
                }
                else
                {
                    ProcessChildren(element, newFormat, elements, ctx, inPre);
                    if (!string.IsNullOrEmpty(href))
                    {
                        var linkText = element.TextContent.Trim();
                        if (linkText != href)
                        {
                            AddTextRun($" ({href})", format, ctx);
                        }
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
            case "abbr" or "acronym":
            {
                ProcessChildren(element, newFormat, elements, ctx, inPre);
                var title = element.GetAttribute("title");
                if (!string.IsNullOrEmpty(title) && ctx.MainPart != null)
                {
                    ctx.CurrentRuns.Add(BuildFootnoteRun(ctx, title!));
                }

                return;
            }
        }

        // Handle blockquote with cite attribute as footnote
        if (tag == "blockquote")
        {
            var cite = element.GetAttribute("cite");
            if (!string.IsNullOrEmpty(cite) && ctx.MainPart != null)
            {
                FlushParagraph(elements, ctx);
                ProcessChildren(element, newFormat, elements, ctx, inPre);
                ctx.CurrentRuns.Add(BuildFootnoteRun(ctx, cite!));
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

                // Paragraph spacing and alignment
                var pf = new ParagraphFormatState();
                if (declarations.TryGetValue("margin", out var marginShorthand))
                {
                    var (t, r, b, l) = StyleParser.ParseMarginShorthand(marginShorthand);
                    pf.MarginTopTwips ??= t;
                    pf.MarginRightTwips ??= r;
                    pf.MarginBottomTwips ??= b;
                    pf.MarginLeftTwips ??= l;
                }

                if (declarations.TryGetValue("margin-top", out var mt))
                {
                    pf.MarginTopTwips = StyleParser.ParseLengthToTwips(mt);
                }

                if (declarations.TryGetValue("margin-bottom", out var mb))
                {
                    pf.MarginBottomTwips = StyleParser.ParseLengthToTwips(mb);
                }

                if (declarations.TryGetValue("margin-left", out var ml))
                {
                    pf.MarginLeftTwips = StyleParser.ParseLengthToTwips(ml);
                }

                if (declarations.TryGetValue("margin-right", out var mr))
                {
                    pf.MarginRightTwips = StyleParser.ParseLengthToTwips(mr);
                }

                if (declarations.TryGetValue("text-indent", out var ti))
                {
                    pf.TextIndentTwips = StyleParser.ParseLengthToTwips(ti);
                }

                if (declarations.TryGetValue("text-align", out var ta))
                {
                    pf.TextAlign = StyleParser.ParseTextAlign(ta);
                }

                if (declarations.TryGetValue("line-height", out var lh))
                {
                    var lhSpan = lh.AsSpan().Trim();
                    if (lhSpan.EndsWith("%".AsSpan()))
                    {
                        if (double.TryParse(lhSpan[..^1].ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct))
                        {
                            pf.LineHeightMultiple = pct / 100.0;
                        }
                    }
                    else if (double.TryParse(lh, NumberStyles.Float, CultureInfo.InvariantCulture, out var multiple) &&
                             !lh.Contains("pt") && !lh.Contains("px") && !lh.Contains("em"))
                    {
                        pf.LineHeightMultiple = multiple;
                    }
                    else
                    {
                        pf.LineHeightTwips = StyleParser.ParseLengthToTwips(lh);
                    }
                }

                if (pf.HasProperties)
                {
                    ctx.ParagraphFormat = pf;
                }
            }

            FlushParagraph(elements, ctx);

            if (tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
            {
                ctx.HeadingLevel = tag[1] - '0';
            }

            // CSS class → Word style mapping
            if (ctx.StyleMap != null && element.ClassList.Length > 0)
            {
                var (paraStyle, runStyle) = WordStyleLookup.LookupClasses(element, ctx.StyleMap);
                if (paraStyle != null)
                {
                    ctx.ParagraphStyleId = paraStyle;
                }

                if (runStyle != null)
                {
                    newFormat.RunStyleId = runStyle;
                }
            }

            if (pageBreakBefore)
            {
                elements.Add(new Paragraph(
                    new ParagraphProperties(new PageBreakBefore())));
            }
        }
        else if (ctx.StyleMap != null && element.ClassList.Length > 0)
        {
            // Inline elements: check for character style
            var (_, runStyle) = WordStyleLookup.LookupClasses(element, ctx.StyleMap);
            if (runStyle != null)
            {
                newFormat.RunStyleId = runStyle;
            }
        }

        // Add bookmark for elements with id attribute
        var elementId = element.GetAttribute("id") ?? element.GetAttribute("name");
        if (elementId != null && isBlock)
        {
            ctx.BookmarkId++;
            ctx.CurrentRuns.Add(new BookmarkStart { Id = ctx.BookmarkId.ToString(), Name = elementId });
        }

        ProcessChildren(element, newFormat, elements, ctx, inPre);

        if (elementId != null && isBlock)
        {
            ctx.CurrentRuns.Add(new BookmarkEnd { Id = ctx.BookmarkId.ToString() });
        }

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
            ctx.ParagraphStyleId = null;
            ctx.ParagraphFormat = null;
            ctx.ListNumId = null;
            ctx.ListIlvl = null;
            return;
        }

        var paragraph = WordHtmlConverter.BuildParagraph(ctx.CurrentRuns, ctx.ListNumId != null ? 0 : ctx.ListDepth);

        // Apply paragraph style: heading > CSS class > default
        if (ctx.HeadingLevel > 0)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId = new() { Val = $"Heading{ctx.HeadingLevel}" };
        }
        else if (ctx.ParagraphStyleId != null)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId = new() { Val = ctx.ParagraphStyleId };
        }

        // Apply real Word numbering
        if (ctx.ListNumId != null)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId ??= new() { Val = "ListParagraph" };
            paragraph.ParagraphProperties.Append(
                new NumberingProperties(
                    new NumberingLevelReference { Val = ctx.ListIlvl ?? 0 },
                    new NumberingId { Val = ctx.ListNumId.Value }));
        }

        // Apply paragraph format (CSS margins, alignment, line-height)
        if (ctx.ParagraphFormat is { HasProperties: true })
        {
            paragraph.ParagraphProperties ??= new();
            ApplyParagraphFormat(paragraph.ParagraphProperties, ctx.ParagraphFormat);
        }

        elements.Add(paragraph);
        ctx.CurrentRuns = [];
        ctx.ListDepth = 0;
        ctx.HeadingLevel = 0;
        ctx.ParagraphStyleId = null;
        ctx.ParagraphFormat = null;
        ctx.ListNumId = null;
        ctx.ListIlvl = null;
    }

    static void ApplyParagraphFormat(ParagraphProperties props, ParagraphFormatState pf)
    {
        if (pf.MarginTopTwips != null || pf.MarginBottomTwips != null ||
            pf.LineHeightMultiple != null || pf.LineHeightTwips != null)
        {
            var spacing = new SpacingBetweenLines();
            if (pf.MarginTopTwips != null)
            {
                spacing.Before = pf.MarginTopTwips.Value.ToString();
            }

            if (pf.MarginBottomTwips != null)
            {
                spacing.After = pf.MarginBottomTwips.Value.ToString();
            }

            if (pf.LineHeightMultiple != null)
            {
                spacing.Line = ((int)(pf.LineHeightMultiple.Value * 240)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            else if (pf.LineHeightTwips != null)
            {
                spacing.Line = pf.LineHeightTwips.Value.ToString();
                spacing.LineRule = LineSpacingRuleValues.Exact;
            }

            props.Append(spacing);
        }

        if (pf.MarginLeftTwips != null || pf.MarginRightTwips != null || pf.TextIndentTwips != null)
        {
            var indent = props.GetFirstChild<Indentation>() ?? new Indentation();
            if (props.GetFirstChild<Indentation>() == null)
            {
                props.Append(indent);
            }

            if (pf.MarginLeftTwips != null)
            {
                indent.Left = pf.MarginLeftTwips.Value.ToString();
            }

            if (pf.MarginRightTwips != null)
            {
                indent.Right = pf.MarginRightTwips.Value.ToString();
            }

            if (pf.TextIndentTwips != null)
            {
                if (pf.TextIndentTwips.Value >= 0)
                {
                    indent.FirstLine = pf.TextIndentTwips.Value.ToString();
                }
                else
                {
                    indent.Hanging = (-pf.TextIndentTwips.Value).ToString();
                }
            }
        }

        if (pf.TextAlign != null)
        {
            props.Append(new Justification { Val = pf.TextAlign.Value });
        }
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

        var tblPr = new TableProperties(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" }));

        var tableStyle = tableElement.GetAttribute("style");
        if (tableStyle != null)
        {
            var declarations = StyleParser.Parse(tableStyle);

            if (declarations.TryGetValue("width", out var tableWidth))
            {
                var twips = StyleParser.ParseLengthToTwips(tableWidth);
                if (twips != null)
                {
                    tblPr.TableWidth = new TableWidth { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa };
                }
            }

            if (declarations.TryGetValue("background-color", out var tableBg) ||
                declarations.TryGetValue("background", out tableBg))
            {
                var parsed = ColorParser.Parse(tableBg);
                if (parsed != null)
                {
                    tblPr.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = parsed });
                }
            }

            ApplyTableCellPadding(declarations, tblPr);
        }

        // HTML cellpadding attribute
        var cellPaddingAttr = tableElement.GetAttribute("cellpadding");
        if (cellPaddingAttr != null)
        {
            var twips = StyleParser.ParseLengthToTwips(cellPaddingAttr);
            if (twips != null)
            {
                var margin = new TableCellMarginDefault(
                    new TopMargin { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa },
                    new StartMargin { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa },
                    new EndMargin { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa });
                tblPr.Append(margin);
            }
        }

        table.Append(tblPr);

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
        TableCellProperties? tcPr = null;

        if (colspan > 1 || isRowspanStart)
        {
            tcPr = new TableCellProperties();
            if (colspan > 1)
            {
                tcPr.Append(new GridSpan { Val = colspan });
            }

            if (isRowspanStart)
            {
                tcPr.Append(new VerticalMerge { Val = MergedCellValues.Restart });
            }
        }

        var cellStyle = cellElement.GetAttribute("style");
        if (cellStyle != null)
        {
            var declarations = StyleParser.Parse(cellStyle);
            tcPr = ApplyCellStyles(declarations, tcPr);
        }

        // HTML bgcolor attribute
        var bgColorAttr = cellElement.GetAttribute("bgcolor");
        if (bgColorAttr != null)
        {
            var parsed = ColorParser.Parse(bgColorAttr);
            if (parsed != null)
            {
                tcPr ??= new TableCellProperties();
                tcPr.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = parsed });
            }
        }

        // HTML width attribute
        var widthAttr = cellElement.GetAttribute("width");
        if (widthAttr != null)
        {
            var twips = StyleParser.ParseLengthToTwips(widthAttr);
            if (twips != null)
            {
                tcPr ??= new TableCellProperties();
                tcPr.Append(new TableCellWidth { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }
        }

        if (tcPr != null)
        {
            tc.Append(tcPr);
        }

        var cellFormat = format.Copy();
        HtmlSegmentParser.ApplyElementFormatting(cellElement, cellElement.LocalName, cellFormat);

        var cellElements = new List<OpenXmlElement>();
        var cellCtx = new WordBuildContext
        {
            MainPart = parentCtx.MainPart,
            ImageIndex = parentCtx.ImageIndex,
            Settings = parentCtx.Settings,
            StyleMap = parentCtx.StyleMap,
            BulletAbstractNumId = parentCtx.BulletAbstractNumId,
            DecimalAbstractNumId = parentCtx.DecimalAbstractNumId,
            NextNumId = parentCtx.NextNumId
        };
        ProcessChildren(cellElement, cellFormat, cellElements, cellCtx, false);
        FlushParagraph(cellElements, cellCtx);
        parentCtx.ImageIndex = cellCtx.ImageIndex;
        parentCtx.NextNumId = cellCtx.NextNumId;
        parentCtx.BulletAbstractNumId = cellCtx.BulletAbstractNumId;
        parentCtx.DecimalAbstractNumId = cellCtx.DecimalAbstractNumId;

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

    static TableCellProperties ApplyCellStyles(Dictionary<string, string> declarations, TableCellProperties? tcPr)
    {
        tcPr ??= new TableCellProperties();

        // Cell width
        if (declarations.TryGetValue("width", out var cellWidth))
        {
            var twips = StyleParser.ParseLengthToTwips(cellWidth);
            if (twips != null)
            {
                tcPr.Append(new TableCellWidth { Width = twips.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }
        }

        // Background color
        if (declarations.TryGetValue("background-color", out var bgColor) ||
            declarations.TryGetValue("background", out bgColor))
        {
            var parsed = ColorParser.Parse(bgColor);
            if (parsed != null)
            {
                tcPr.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = parsed });
            }
        }

        // Vertical alignment
        if (declarations.TryGetValue("vertical-align", out var vAlign))
        {
            var val = vAlign.Trim().ToLowerInvariant() switch
            {
                "top" => TableVerticalAlignmentValues.Top,
                "middle" => TableVerticalAlignmentValues.Center,
                "bottom" => TableVerticalAlignmentValues.Bottom,
                _ => (TableVerticalAlignmentValues?)null
            };
            if (val != null)
            {
                tcPr.Append(new TableCellVerticalAlignment { Val = val.Value });
            }
        }

        // Cell padding
        var hasPadding = false;
        int? padTop = null, padRight = null, padBottom = null, padLeft = null;

        if (declarations.TryGetValue("padding", out var padShorthand))
        {
            var (t, r, b, l) = StyleParser.ParseMarginShorthand(padShorthand);
            padTop = t;
            padRight = r;
            padBottom = b;
            padLeft = l;
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-top", out var pt))
        {
            padTop = StyleParser.ParseLengthToTwips(pt);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-right", out var pr))
        {
            padRight = StyleParser.ParseLengthToTwips(pr);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-bottom", out var pb))
        {
            padBottom = StyleParser.ParseLengthToTwips(pb);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-left", out var pl))
        {
            padLeft = StyleParser.ParseLengthToTwips(pl);
            hasPadding = true;
        }

        if (hasPadding)
        {
            var margin = new TableCellMargin();
            if (padTop != null)
            {
                margin.Append(new TopMargin { Width = padTop.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            if (padLeft != null)
            {
                margin.Append(new StartMargin { Width = padLeft.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            if (padBottom != null)
            {
                margin.Append(new BottomMargin { Width = padBottom.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            if (padRight != null)
            {
                margin.Append(new EndMargin { Width = padRight.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            tcPr.Append(margin);
        }

        return tcPr;
    }

    static void ApplyTableCellPadding(Dictionary<string, string> declarations, TableProperties tblPr)
    {
        var hasPadding = false;
        int? padTop = null, padRight = null, padBottom = null, padLeft = null;

        if (declarations.TryGetValue("padding", out var padShorthand))
        {
            var (t, r, b, l) = StyleParser.ParseMarginShorthand(padShorthand);
            padTop = t;
            padRight = r;
            padBottom = b;
            padLeft = l;
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-top", out var pt))
        {
            padTop = StyleParser.ParseLengthToTwips(pt);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-right", out var pr))
        {
            padRight = StyleParser.ParseLengthToTwips(pr);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-bottom", out var pb))
        {
            padBottom = StyleParser.ParseLengthToTwips(pb);
            hasPadding = true;
        }

        if (declarations.TryGetValue("padding-left", out var pl))
        {
            padLeft = StyleParser.ParseLengthToTwips(pl);
            hasPadding = true;
        }

        if (hasPadding)
        {
            var margin = new TableCellMarginDefault();
            if (padTop != null)
            {
                margin.Append(new TopMargin { Width = padTop.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            if (padLeft != null)
            {
                margin.Append(new StartMargin { Width = padLeft.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            if (padBottom != null)
            {
                margin.Append(new BottomMargin { Width = padBottom.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            if (padRight != null)
            {
                margin.Append(new EndMargin { Width = padRight.Value.ToString(), Type = TableWidthUnitValues.Dxa });
            }

            tblPr.Append(margin);
        }
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

    static Run BuildFootnoteRun(WordBuildContext ctx, string footnoteText)
    {
        ctx.FootnoteIndex++;
        var footnoteId = ctx.FootnoteIndex;

        var footnotesPart = ctx.MainPart!.FootnotesPart;
        if (footnotesPart == null)
        {
            footnotesPart = ctx.MainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(
                    new Paragraph(
                        new Run(
                            new SeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.Separator,
                    Id = -1
                },
                new Footnote(
                    new Paragraph(
                        new Run(
                            new ContinuationSeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.ContinuationSeparator,
                    Id = 0
                });
        }

        footnotesPart.Footnotes!.Append(
            new Footnote(
                new Paragraph(
                    new Run(
                        new RunProperties(
                            new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                        new FootnoteReferenceMark()),
                    new Run(
                        new Text(" " + footnoteText) { Space = SpaceProcessingModeValues.Preserve })))
            {
                Id = footnoteId
            });

        var run = new Run(
            new RunProperties(
                new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new FootnoteReference { Id = footnoteId });

        return run;
    }
}
