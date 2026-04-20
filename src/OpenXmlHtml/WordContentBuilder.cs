static partial class WordContentBuilder
{
    static readonly HtmlParser parser = new();

    internal static List<OpenXmlElement> Build(string html, MainDocumentPart? main, HtmlConvertSettings? settings = null)
    {
        var document = parser.ParseDocument(string.Concat("<body>", html, "</body>"));
        var body = document.Body!;
        var elements = new List<OpenXmlElement>();
        var context = new WordBuildContext
        {
            MainPart = main,
            Settings = settings,
            StyleMap = WordStyleLookup.BuildStyleMap(main)
        };
        if (main?.NumberingDefinitionsPart?.Numbering is { } existingNumbering)
        {
            context.NextNumId = WordNumberingBuilder.GetNextId(existingNumbering);
        }
        else
        {
            context.NextNumId = 1;
        }

        ProcessChildren(body, new(), elements, context, false);
        FlushParagraph(elements, context);
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
                if (imageData == null)
                {
                    var alt = element.GetAttribute("alt");
                    if (!string.IsNullOrEmpty(alt))
                    {
                        // ReSharper disable once RedundantSuppressNullableWarningExpression
                        AddTextRun(alt!, format, context);
                    }
                }
                else
                {
                    if (context.MainPart != null)
                    {
                        context.ImageIndex++;
                        context.CurrentRuns.Add(WordHtmlConverter.BuildImageRun(context.MainPart, imageData, context.ImageIndex));
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
                BuildList(element, tag, newFormat, elements, context, inPre);
                return;
            case "li":
                BuildListItem(element, newFormat, elements, context, inPre);
                return;
            case "a":
                BuildAnchor(element, format, newFormat, elements, context, inPre);
                return;
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
                    // ReSharper disable once RedundantSuppressNullableWarningExpression
                    context.CurrentRuns.Add(BuildFootnoteRun(context, title!));
                }

                return;
            }
        }

        // Handle blockquote with cite attribute as footnote
        if (tag == "blockquote")
        {
            var cite = element.GetAttribute("cite");
            if (!string.IsNullOrEmpty(cite) &&
                context.MainPart != null)
            {
                FlushParagraph(elements, context);
                ProcessChildren(element, newFormat, elements, context, inPre);
                // ReSharper disable once RedundantSuppressNullableWarningExpression
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
                pageBreakBefore = styleDeclarations.TryGetValue("page-break-before", out var pbb) && pbb == "always";
                pageBreakAfter = styleDeclarations.TryGetValue("page-break-after", out var pba) && pba == "always";

                var pf = ParagraphFormatState.ParseFrom(styleDeclarations);
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
            if (context.StyleMap != null &&
                element.ClassList.Length > 0)
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
                elements.Add(
                    new Paragraph(
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
            context.CurrentRuns.Add(
                new BookmarkStart
                {
                    Id = bookmarkId,
                    Name = elementId
                });
        }

        ProcessChildren(element, newFormat, elements, context, inPre);

        if (bookmarkId != null)
        {
            context.CurrentRuns.Add(
                new BookmarkEnd
                {
                    Id = bookmarkId
                });
        }

        if (isBlock)
        {
            FlushParagraph(elements, context);

            if (pageBreakAfter)
            {
                elements.Add(
                    new Paragraph(
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
        XmlCharFilter.StripInvalidXmlChars(
            transform switch
            {
                "uppercase" => text.ToUpperInvariant(),
                "lowercase" => text.ToLowerInvariant(),
                "capitalize" => CapitalizeWords(text),
                _ => text
            });

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

    static void FlushParagraph(List<OpenXmlElement> elements, WordBuildContext context)
    {
        if (context.CurrentRuns.Count == 0)
        {
            context.HeadingLevel = 0;
            context.ParagraphStyleId = null;
            context.ParagraphFormat = null;
            context.ListNumId = null;
            context.ListIlvl = null;
            return;
        }

        var paragraph = WordHtmlConverter.BuildParagraph(context.CurrentRuns, context.ListNumId != null ? 0 : context.ListDepth);

        // Apply paragraph style: heading > CSS class > default
        if (context.HeadingLevel > 0)
        {
            var offset = context.Settings?.HeadingLevelOffset ?? 0;
            var level = Math.Clamp(context.HeadingLevel + offset, 1, 9);
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId = new()
            {
                Val = $"Heading{level}"
            };
        }
        else if (context.ParagraphStyleId != null)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId = new()
            {
                Val = context.ParagraphStyleId
            };
        }

        // Apply real Word numbering
        if (context.ListNumId != null)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.ParagraphStyleId ??= new()
            {
                Val = "ListParagraph"
            };
            paragraph.ParagraphProperties.Append(
                new NumberingProperties(
                    new NumberingLevelReference
                    {
                        Val = context.ListIlvl ?? 0
                    },
                    new NumberingId
                    {
                        Val = context.ListNumId.Value
                    }));
        }

        // Apply paragraph format (CSS margins, alignment, line-height)
        if (context.ParagraphFormat is { HasProperties: true })
        {
            paragraph.ParagraphProperties ??= new();
            ApplyParagraphFormat(paragraph.ParagraphProperties, context.ParagraphFormat);
        }

        elements.Add(paragraph);
        context.CurrentRuns.Clear();
        context.ListDepth = 0;
        context.HeadingLevel = 0;
        context.ParagraphStyleId = null;
        context.ParagraphFormat = null;
        context.ListNumId = null;
        context.ListIlvl = null;
    }

    static void ApplyParagraphFormat(ParagraphProperties props, ParagraphFormatState pf)
    {
        if (pf.MarginTopTwips != null ||
            pf.MarginBottomTwips != null ||
            pf.LineHeightMultiple != null ||
            pf.LineHeightTwips != null)
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

        if (pf.MarginLeftTwips != null ||
            pf.MarginRightTwips != null ||
            pf.TextIndentTwips != null)
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
            props.Append(
                new Justification
                {
                    Val = pf.TextAlign.Value
                });
        }

        if (pf.BackgroundColor != null)
        {
            props.Append(
                new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Fill = pf.BackgroundColor
                });
        }

        if (pf.WritingMode != null)
        {
            props.Append(new BiDi());
            props.Append(
                new TextDirection
                {
                    Val = pf.WritingMode.Value
                });
        }

        if (pf.BorderTop != null ||
            pf.BorderRight != null ||
            pf.BorderBottom != null ||
            pf.BorderLeft != null)
        {
            var borders = new ParagraphBorders();
            BorderEmitter.AppendSides(borders, pf.BorderTop, pf.BorderLeft, pf.BorderBottom, pf.BorderRight, 1);
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
        while (elements.Count > 0 &&
               elements[^1] is Paragraph { HasChildren: false })
        {
            elements.RemoveAt(elements.Count - 1);
        }
    }

    static void BuildAnchor(IElement element, FormatState format, FormatState newFormat, List<OpenXmlElement> elements, WordBuildContext context, bool inPre)
    {
        var href = element.GetAttribute("href");

        if (href != null &&
            href.StartsWith('#') &&
            href.Length > 1)
        {
            WrapChildrenInHyperlink(
                element,
                newFormat,
                elements,
                context,
                inPre,
                new()
                {
                    Anchor = href[1..]
                });
            return;
        }

        if (!string.IsNullOrEmpty(href) && context.MainPart != null &&
            Uri.TryCreate(href, UriKind.Absolute, out var uri))
        {
            var rel = context.MainPart.AddHyperlinkRelationship(uri, true);
            WrapChildrenInHyperlink(
                element,
                newFormat,
                elements,
                context,
                inPre,
                new()
                {
                    Id = rel.Id
                });
            return;
        }

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

    static void WrapChildrenInHyperlink(IElement element, FormatState newFormat, List<OpenXmlElement> elements, WordBuildContext context, bool inPre, Hyperlink hyperlink)
    {
        var runsBefore = context.CurrentRuns.Count;
        ProcessChildren(element, newFormat, elements, context, inPre);
        while (context.CurrentRuns.Count > runsBefore)
        {
            var run = context.CurrentRuns[runsBefore];
            context.CurrentRuns.RemoveAt(runsBefore);
            hyperlink.Append(run);
        }

        context.CurrentRuns.Add(hyperlink);
    }
}
