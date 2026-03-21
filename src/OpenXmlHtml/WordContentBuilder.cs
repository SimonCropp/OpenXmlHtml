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
    internal int NextNumId;
    internal int? ListNumId;
    internal int? ListIlvl;
    internal int? ReversedStart;
}

static class WordContentBuilder
{
    static readonly HtmlParser parser = new();

    internal static List<OpenXmlElement> Build(string html, MainDocumentPart? mainPart, HtmlConvertSettings? settings = null)
    {
        var document = parser.ParseDocument(string.Concat("<body>", html, "</body>"));
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

    static void ProcessElement(IElement element, FormatState format, List<OpenXmlElement> elements, WordBuildContext context, bool inPre)
    {
        var tag = element.LocalName;
        var newFormat = format;
        HtmlSegmentParser.ApplyElementFormatting(element, tag, ref newFormat, out var styleDeclarations);

        switch (tag)
        {
            case "br":
                ForceFlushParagraph(elements, context);
                return;
            case "hr":
                FlushParagraph(elements, context);
                AddTextRun("\u2014\u2014\u2014", format, context);
                FlushParagraph(elements, context);
                return;
            case "img":
            {
                var imageData = ImageResolver.Resolve(element, context.Settings);
                if (imageData != null)
                {
                    if (context.MainPart != null)
                    {
                        context.ImageIndex++;
                        context.CurrentRuns.Add(WordHtmlConverter.BuildImageRun(context.MainPart, imageData, context.ImageIndex));
                    }
                }
                else
                {
                    var alt = element.GetAttribute("alt");
                    if (!string.IsNullOrEmpty(alt))
                    {
                        // ReSharper disable once RedundantSuppressNullableWarningExpression
                        AddTextRun(alt!, format, context);
                    }
                }

                return;
            }
            case "svg":
            {
                if (context.MainPart != null)
                {
                    var imageData = HtmlSegmentParser.ParseSvgElement(element);
                    context.ImageIndex++;
                    context.CurrentRuns.Add(WordHtmlConverter.BuildImageRun(context.MainPart, imageData, context.ImageIndex));
                }

                return;
            }
            case "col":
                return;
            case "table":
                FlushParagraph(elements, context);
                BuildTable(element, format, elements, context);
                return;
            case "ul" or "ol":
            {
                FlushParagraph(elements, context);
                var isReversed = tag == "ol" && element.HasAttribute("reversed");
                if (context.MainPart != null && !isReversed)
                {
                    var isOrdered = tag == "ol";
                    var part = WordNumberingBuilder.EnsureNumberingPart(context.MainPart);
                    var numbering = part.Numbering!;

                    // Determine list format from type attribute or list-style-type CSS
                    var typeAttr = element.GetAttribute("type");
                    var listStyleCss = element.GetAttribute("style") is { } listStyle
                        ? StyleParser.Parse(listStyle).GetValueOrDefault("list-style-type")
                        : null;
                    var format2 = WordNumberingBuilder.ParseListStyleType(typeAttr, listStyleCss, isOrdered);

                    int abstractNumId;
                    if (format2 == NumberFormatValues.Bullet)
                    {
                        if (context.BulletAbstractNumId == null)
                        {
                            var id = WordNumberingBuilder.GetNextId(numbering);
                            context.BulletAbstractNumId = WordNumberingBuilder.CreateBulletAbstractNum(numbering, id);
                        }

                        abstractNumId = context.BulletAbstractNumId.Value;
                    }
                    else
                    {
                        // Each ordered format type gets its own abstract num
                        var id = WordNumberingBuilder.GetNextId(numbering);
                        abstractNumId = WordNumberingBuilder.CreateOrderedAbstractNum(numbering, id, format2);
                    }

                    // Parse start attribute for ordered lists
                    int? startOverride = null;
                    var startAttr = element.GetAttribute("start");
                    if (startAttr != null && int.TryParse(startAttr, out var startVal) && startVal != 1)
                    {
                        startOverride = startVal;
                    }

                    var numId = WordNumberingBuilder.GetNextId(numbering);
                    WordNumberingBuilder.AddNumberingInstance(numbering, numId, abstractNumId, startOverride);
                    var ilvl = context.ListStack.Count > 0 ? context.ListStack.Peek().Ilvl + 1 : 0;
                    context.ListStack.Push((numId, ilvl, isOrdered));
                    ProcessChildren(element, newFormat, elements, context, inPre);
                    context.ListStack.Pop();
                }
                else
                {
                    if (isReversed)
                    {
                        // Count <li> children for reversed numbering
                        var liCount = 0;
                        foreach (var child in element.Children)
                        {
                            if (child.LocalName == "li")
                            {
                                liCount++;
                            }
                        }

                        var startAttr2 = element.GetAttribute("start");
                        context.ReversedStart = startAttr2 != null && int.TryParse(startAttr2, out var s) ? s : liCount;
                    }

                    ProcessChildren(element, newFormat, elements, context, inPre);
                    context.ReversedStart = null;
                }

                FlushParagraph(elements, context);
                return;
            }
            case "li":
            {
                FlushParagraph(elements, context);
                if (context.ListStack.Count > 0)
                {
                    var (numId, ilvl, _) = context.ListStack.Peek();
                    context.ListNumId = numId;
                    context.ListIlvl = ilvl;
                }
                else
                {
                    // Fallback: text prefix when no MainDocumentPart
                    var parent = element.ParentElement?.LocalName;
                    var depth = HtmlSegmentParser.GetListDepth(element);
                    context.ListDepth = depth;

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

                        if (context.ReversedStart != null)
                        {
                            // Reversed: count down from start
                            index = context.ReversedStart.Value - (index - 1);
                        }

                        AddTextRun($"{index}. ", newFormat, context);
                    }
                    else
                    {
                        var bullet = depth switch
                        {
                            0 => "\u25CF",
                            1 => "\u25CB",
                            _ => "\u25A0"
                        };
                        AddTextRun($"{bullet} ", newFormat, context);
                    }
                }

                ProcessChildren(element, newFormat, elements, context, inPre);
                FlushParagraph(elements, context);
                return;
            }
            case "a":
            {
                var href = element.GetAttribute("href");
                if (href != null && href.StartsWith('#') && href.Length > 1)
                {
                    // Internal anchor link → Word hyperlink to bookmark
                    var runsBefore = context.CurrentRuns.Count;
                    ProcessChildren(element, newFormat, elements, context, inPre);
                    var hyperlink = new Hyperlink { Anchor = href[1..] };
                    // Move newly added runs into the hyperlink
                    while (context.CurrentRuns.Count > runsBefore)
                    {
                        var run = context.CurrentRuns[runsBefore];
                        context.CurrentRuns.RemoveAt(runsBefore);
                        hyperlink.Append(run);
                    }

                    context.CurrentRuns.Add(hyperlink);
                }
                else if (!string.IsNullOrEmpty(href) && context.MainPart != null &&
                         Uri.TryCreate(href, UriKind.Absolute, out var uri))
                {
                    // External hyperlink with relationship
                    var runsBefore = context.CurrentRuns.Count;
                    ProcessChildren(element, newFormat, elements, context, inPre);
                    var rel = context.MainPart.AddHyperlinkRelationship(uri, true);
                    var hyperlink = new Hyperlink { Id = rel.Id };
                    while (context.CurrentRuns.Count > runsBefore)
                    {
                        var run = context.CurrentRuns[runsBefore];
                        context.CurrentRuns.RemoveAt(runsBefore);
                        hyperlink.Append(run);
                    }

                    context.CurrentRuns.Add(hyperlink);
                }
                else
                {
                    ProcessChildren(element, newFormat, elements, context, inPre);
                    if (!string.IsNullOrEmpty(href))
                    {
                        var linkText = element.TextContent.Trim();
                        if (linkText != href)
                        {
                            AddTextRun($" ({href})", format, context);
                        }
                    }
                }

                return;
            }
            case "q":
            {
                AddTextRun("\u201C", format, context);
                ProcessChildren(element, newFormat, elements, context, inPre);
                AddTextRun("\u201D", format, context);
                return;
            }
            case "pre":
            {
                FlushParagraph(elements, context);
                ProcessChildren(element, newFormat, elements, context, true);
                FlushParagraph(elements, context);
                return;
            }
            case "abbr" or "acronym":
            {
                ProcessChildren(element, newFormat, elements, context, inPre);
                var title = element.GetAttribute("title");
                if (!string.IsNullOrEmpty(title) &&
                    context.MainPart != null)
                {
                    context.CurrentRuns.Add(BuildFootnoteRun(context, title!));
                }

                return;
            }
        }

        // Handle blockquote with cite attribute as footnote
        if (tag == "blockquote")
        {
            var cite = element.GetAttribute("cite");
            if (!string.IsNullOrEmpty(cite) && context.MainPart != null)
            {
                FlushParagraph(elements, context);
                ProcessChildren(element, newFormat, elements, context, inPre);
                context.CurrentRuns.Add(BuildFootnoteRun(context, cite!));
                FlushParagraph(elements, context);
                return;
            }
        }

        var isBlock = HtmlSegmentParser.IsBlockElement(tag);
        var pageBreakBefore = false;
        var pageBreakAfter = false;

        if (isBlock)
        {
            if (styleDeclarations != null)
            {
                var declarations = styleDeclarations;
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

                if (declarations.TryGetValue("background-color", out var blockBg) ||
                    declarations.TryGetValue("background", out blockBg))
                {
                    var parsed = ColorParser.Parse(blockBg);
                    if (parsed != null)
                    {
                        pf.BackgroundColor = parsed;
                    }
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

                if (declarations.TryGetValue("writing-mode", out var writingMode))
                {
                    var wm = writingMode.AsSpan().Trim();
                    pf.WritingMode = wm.Equals("vertical-rl", StringComparison.OrdinalIgnoreCase) || wm.Equals("tb-rl", StringComparison.OrdinalIgnoreCase)
                        ? TextDirectionValues.TopToBottomRightToLeft
                        : wm.Equals("vertical-lr", StringComparison.OrdinalIgnoreCase) || wm.Equals("tb-lr", StringComparison.OrdinalIgnoreCase)
                        ? TextDirectionValues.BottomToTopLeftToRight
                        : null;
                }

                if (declarations.TryGetValue("direction", out var direction) &&
                    direction.Trim().Equals("rtl", StringComparison.OrdinalIgnoreCase))
                {
                    pf.WritingMode ??= TextDirectionValues.TopToBottomRightToLeft;
                }

                if (declarations.TryGetValue("border", out var borderAll))
                {
                    var bi = StyleParser.ParseBorder(borderAll);
                    pf.BorderTop ??= bi;
                    pf.BorderRight ??= bi;
                    pf.BorderBottom ??= bi;
                    pf.BorderLeft ??= bi;
                }

                if (declarations.TryGetValue("border-top", out var bt))
                {
                    pf.BorderTop = StyleParser.ParseBorder(bt);
                }

                if (declarations.TryGetValue("border-right", out var br2))
                {
                    pf.BorderRight = StyleParser.ParseBorder(br2);
                }

                if (declarations.TryGetValue("border-bottom", out var bb))
                {
                    pf.BorderBottom = StyleParser.ParseBorder(bb);
                }

                if (declarations.TryGetValue("border-left", out var bl))
                {
                    pf.BorderLeft = StyleParser.ParseBorder(bl);
                }

                if (pf.HasProperties)
                {
                    context.ParagraphFormat = pf;
                }
            }

            FlushParagraph(elements, context);

            if (tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
            {
                context.HeadingLevel = tag[1] - '0';
            }

            // CSS class → Word style mapping
            if (context.StyleMap != null && element.ClassList.Length > 0)
            {
                var (paraStyle, runStyle) = WordStyleLookup.LookupClasses(element, context.StyleMap);
                if (paraStyle != null)
                {
                    context.ParagraphStyleId = paraStyle;
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
        else if (context.StyleMap != null && element.ClassList.Length > 0)
        {
            // Inline elements: check for character style
            var (_, runStyle) = WordStyleLookup.LookupClasses(element, context.StyleMap);
            if (runStyle != null)
            {
                newFormat.RunStyleId = runStyle;
            }
        }

        // Add bookmark for elements with id attribute
        var elementId = element.GetAttribute("id") ?? element.GetAttribute("name");
        string? bookmarkId = null;
        if (elementId != null && isBlock)
        {
            context.BookmarkId++;
            bookmarkId = context.BookmarkId.ToString();
            context.CurrentRuns.Add(new BookmarkStart { Id = bookmarkId, Name = elementId });
        }

        ProcessChildren(element, newFormat, elements, context, inPre);

        if (bookmarkId != null)
        {
            context.CurrentRuns.Add(new BookmarkEnd { Id = bookmarkId });
        }

        if (isBlock)
        {
            FlushParagraph(elements, context);

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
            new Text(ApplyTextTransform(text, format.TextTransform))
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        ctx.CurrentRuns.Add(run);
    }

    internal static string ApplyTextTransform(string text, string? transform) =>
        transform switch
        {
            "uppercase" => text.ToUpperInvariant(),
            "lowercase" => text.ToLowerInvariant(),
            "capitalize" => CapitalizeWords(text),
            _ => text
        };

    static string CapitalizeWords(string text)
    {
        var chars = text.ToCharArray();
        var capitalizeNext = true;
        for (var i = 0; i < chars.Length; i++)
        {
            if (char.IsWhiteSpace(chars[i]))
            {
                capitalizeNext = true;
            }
            else if (capitalizeNext)
            {
                chars[i] = char.ToUpperInvariant(chars[i]);
                capitalizeNext = false;
            }
        }

        return new(chars);
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
        ctx.CurrentRuns.Clear();
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

        if (pf.BackgroundColor != null)
        {
            props.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = pf.BackgroundColor });
        }

        if (pf.WritingMode != null)
        {
            props.Append(new BiDi());
            props.Append(new TextDirection { Val = pf.WritingMode.Value });
        }

        if (pf.BorderTop != null || pf.BorderRight != null || pf.BorderBottom != null || pf.BorderLeft != null)
        {
            var borders = new ParagraphBorders();
            if (pf.BorderTop != null && pf.BorderTop.Style != BorderValues.None)
            {
                borders.Append(new TopBorder { Val = pf.BorderTop.Style, Size = (uint)pf.BorderTop.SizeEighths, Space = 1, Color = pf.BorderTop.Color ?? "auto" });
            }

            if (pf.BorderLeft != null && pf.BorderLeft.Style != BorderValues.None)
            {
                borders.Append(new LeftBorder { Val = pf.BorderLeft.Style, Size = (uint)pf.BorderLeft.SizeEighths, Space = 1, Color = pf.BorderLeft.Color ?? "auto" });
            }

            if (pf.BorderBottom != null && pf.BorderBottom.Style != BorderValues.None)
            {
                borders.Append(new BottomBorder { Val = pf.BorderBottom.Style, Size = (uint)pf.BorderBottom.SizeEighths, Space = 1, Color = pf.BorderBottom.Color ?? "auto" });
            }

            if (pf.BorderRight != null && pf.BorderRight.Style != BorderValues.None)
            {
                borders.Append(new RightBorder { Val = pf.BorderRight.Style, Size = (uint)pf.BorderRight.SizeEighths, Space = 1, Color = pf.BorderRight.Color ?? "auto" });
            }

            props.Append(borders);
        }
    }

    static void ForceFlushParagraph(List<OpenXmlElement> elements, WordBuildContext ctx)
    {
        elements.Add(WordHtmlConverter.BuildParagraph(ctx.CurrentRuns, ctx.ListDepth));
        ctx.CurrentRuns.Clear();
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
            var captionFormat = format;
            HtmlSegmentParser.ApplyElementFormatting(caption, "caption", ref captionFormat);
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

        // Default border: single thin line
        var defaultBorder = new BorderInfo(4, BorderValues.Single, "auto");

        // HTML border attribute: border="0" removes borders
        var borderAttr = tableElement.GetAttribute("border");
        if (borderAttr == "0")
        {
            defaultBorder = new(0, BorderValues.None, null);
        }

        var tableStyle = tableElement.GetAttribute("style");
        Dictionary<string, string>? declarations = null;
        if (tableStyle != null)
        {
            declarations = StyleParser.Parse(tableStyle);

            if (declarations.TryGetValue("border", out var tableBorderCss))
            {
                defaultBorder = StyleParser.ParseBorder(tableBorderCss) ?? defaultBorder;
            }
        }

        TableBorders tableBorders;
        if (defaultBorder.Style == BorderValues.None)
        {
            tableBorders = new TableBorders(
                new TopBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                new LeftBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                new BottomBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                new RightBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                new InsideHorizontalBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                new InsideVerticalBorder { Val = BorderValues.None, Size = 0, Space = 0 });
        }
        else
        {
            tableBorders = new TableBorders(
                new TopBorder { Val = defaultBorder.Style, Size = (uint)defaultBorder.SizeEighths, Space = 0, Color = defaultBorder.Color ?? "auto" },
                new LeftBorder { Val = defaultBorder.Style, Size = (uint)defaultBorder.SizeEighths, Space = 0, Color = defaultBorder.Color ?? "auto" },
                new BottomBorder { Val = defaultBorder.Style, Size = (uint)defaultBorder.SizeEighths, Space = 0, Color = defaultBorder.Color ?? "auto" },
                new RightBorder { Val = defaultBorder.Style, Size = (uint)defaultBorder.SizeEighths, Space = 0, Color = defaultBorder.Color ?? "auto" },
                new InsideHorizontalBorder { Val = defaultBorder.Style, Size = (uint)defaultBorder.SizeEighths, Space = 0, Color = defaultBorder.Color ?? "auto" },
                new InsideVerticalBorder { Val = defaultBorder.Style, Size = (uint)defaultBorder.SizeEighths, Space = 0, Color = defaultBorder.Color ?? "auto" });
        }

        var tblPr = new TableProperties(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
            tableBorders);

        if (declarations != null)
        {
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
                var w = twips.Value.ToString();
                var margin = new TableCellMarginDefault(
                    new TopMargin { Width = w, Type = TableWidthUnitValues.Dxa },
                    new StartMargin { Width = w, Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = w, Type = TableWidthUnitValues.Dxa },
                    new EndMargin { Width = w, Type = TableWidthUnitValues.Dxa });
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

            // Row height from style or height attribute
            var rowHeight = row.GetAttribute("style") is { } rowStyle
                ? StyleParser.ParseLengthToTwips(StyleParser.Parse(rowStyle).GetValueOrDefault("height") ?? "")
                : null;
            rowHeight ??= row.GetAttribute("height") is { } rh ? StyleParser.ParseLengthToTwips(rh) : null;
            if (rowHeight != null)
            {
                tr.Append(new TableRowProperties(
                    new TableRowHeight { Val = (uint)rowHeight.Value, HeightType = HeightRuleValues.AtLeast }));
            }

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

        var cellFormat = format;
        HtmlSegmentParser.ApplyElementFormatting(cellElement, cellElement.LocalName, ref cellFormat);

        var cellElements = new List<OpenXmlElement>();
        var cellCtx = new WordBuildContext
        {
            MainPart = parentCtx.MainPart,
            ImageIndex = parentCtx.ImageIndex,
            Settings = parentCtx.Settings,
            StyleMap = parentCtx.StyleMap,
            BulletAbstractNumId = parentCtx.BulletAbstractNumId,
            NextNumId = parentCtx.NextNumId
        };
        ProcessChildren(cellElement, cellFormat, cellElements, cellCtx, false);
        FlushParagraph(cellElements, cellCtx);
        parentCtx.ImageIndex = cellCtx.ImageIndex;
        parentCtx.NextNumId = cellCtx.NextNumId;
        parentCtx.BulletAbstractNumId = cellCtx.BulletAbstractNumId;

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

            // OOXML requires every cell to end with a paragraph
            if (cellElements[^1] is not Paragraph)
            {
                tc.Append(new Paragraph());
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
            var va = vAlign.AsSpan().Trim();
            var val = va.Equals("top", StringComparison.OrdinalIgnoreCase) ? TableVerticalAlignmentValues.Top
                : va.Equals("middle", StringComparison.OrdinalIgnoreCase) ? TableVerticalAlignmentValues.Center
                : va.Equals("bottom", StringComparison.OrdinalIgnoreCase) ? TableVerticalAlignmentValues.Bottom
                : (TableVerticalAlignmentValues?)null;
            if (val != null)
            {
                tcPr.Append(new TableCellVerticalAlignment { Val = val.Value });
            }
        }

        // Writing mode on cell
        if (declarations.TryGetValue("writing-mode", out var cellWritingMode))
        {
            var cwm = cellWritingMode.AsSpan().Trim();
            var cellTextDir = cwm.Equals("vertical-rl", StringComparison.OrdinalIgnoreCase) || cwm.Equals("tb-rl", StringComparison.OrdinalIgnoreCase)
                ? TextDirectionValues.TopToBottomRightToLeft
                : cwm.Equals("vertical-lr", StringComparison.OrdinalIgnoreCase) || cwm.Equals("tb-lr", StringComparison.OrdinalIgnoreCase)
                ? TextDirectionValues.BottomToTopLeftToRight
                : (TextDirectionValues?)null;
            if (cellTextDir != null)
            {
                tcPr.Append(new TextDirection { Val = cellTextDir.Value });
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

        // Cell borders
        BorderInfo? cellBorderAll = null;
        if (declarations.TryGetValue("border", out var cellBorderVal))
        {
            cellBorderAll = StyleParser.ParseBorder(cellBorderVal);
        }

        var cbt = declarations.TryGetValue("border-top", out var cbtVal) ? StyleParser.ParseBorder(cbtVal) : cellBorderAll;
        var cbr = declarations.TryGetValue("border-right", out var cbrVal) ? StyleParser.ParseBorder(cbrVal) : cellBorderAll;
        var cbb = declarations.TryGetValue("border-bottom", out var cbbVal) ? StyleParser.ParseBorder(cbbVal) : cellBorderAll;
        var cbl = declarations.TryGetValue("border-left", out var cblVal) ? StyleParser.ParseBorder(cblVal) : cellBorderAll;

        if (cbt != null || cbr != null || cbb != null || cbl != null)
        {
            var cellBorders = new TableCellBorders();
            if (cbt != null && cbt.Style != BorderValues.None)
            {
                cellBorders.Append(new TopBorder { Val = cbt.Style, Size = (uint)cbt.SizeEighths, Space = 0, Color = cbt.Color ?? "auto" });
            }

            if (cbl != null && cbl.Style != BorderValues.None)
            {
                cellBorders.Append(new LeftBorder { Val = cbl.Style, Size = (uint)cbl.SizeEighths, Space = 0, Color = cbl.Color ?? "auto" });
            }

            if (cbb != null && cbb.Style != BorderValues.None)
            {
                cellBorders.Append(new BottomBorder { Val = cbb.Style, Size = (uint)cbb.SizeEighths, Space = 0, Color = cbb.Color ?? "auto" });
            }

            if (cbr != null && cbr.Style != BorderValues.None)
            {
                cellBorders.Append(new RightBorder { Val = cbr.Style, Size = (uint)cbr.SizeEighths, Space = 0, Color = cbr.Color ?? "auto" });
            }

            tcPr.Append(cellBorders);
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
