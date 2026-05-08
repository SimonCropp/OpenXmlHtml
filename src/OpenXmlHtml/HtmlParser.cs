static class HtmlSegmentParser
{
    static readonly HtmlParser parser = new();

    internal static List<TextSegment> Parse(string html) =>
        Parse(html, null);

    internal static List<TextSegment> Parse(string html, HtmlConvertSettings? settings)
    {
        var document = parser.ParseDocument(string.Concat("<body>", html, "</body>"));
        var body = document.Body!;
        var segments = new List<TextSegment>();
        var format = new FormatState();
        ProcessNode(body, format, segments, false, settings);
        TrimTrailingNewlines(segments);
        return segments;
    }

    static void ProcessNode(INode node, FormatState format, List<TextSegment> segments, bool inPre, HtmlConvertSettings? settings)
    {
        foreach (var child in node.ChildNodes)
        {
            switch (child)
            {
                case IText textNode:
                {
                    var text = inPre ? textNode.Data : CollapseWhitespace(textNode.Data);
                    if (text.Length > 0 &&
                        !(string.IsNullOrWhiteSpace(text) &&
                          IsInterBlockWhitespace(textNode)))
                    {
                        segments.Add(new(text, format));
                    }

                    break;
                }
                case IElement element:
                    ProcessElement(element, format, segments, inPre, settings);
                    break;
            }
        }
    }

    static void ProcessElement(IElement element, FormatState format, List<TextSegment> segments, bool inPre, HtmlConvertSettings? settings)
    {
        var tag = element.LocalName;
        var newFormat = format;
        ApplyElementFormatting(element, tag, ref newFormat, out var styleDeclarations);
        inPre = ApplyWhiteSpace(styleDeclarations, ref newFormat, inPre);

        switch (tag)
        {
            case "br":
                segments.Add(new("\n", format));
                return;
            case "wbr":
                segments.Add(new("\u200B", format));
                return;
            case "hr":
                EnsureNewline(segments, format);
                segments.Add(new("———", format));
                segments.Add(new("\n", format));
                return;
            case "img":
            {
                var imageData = ImageResolver.Resolve(element, settings);
                if (imageData != null)
                {
                    var imageFormat = format;
                    imageFormat.Image = imageData;
                    segments.Add(new("\uFFFC", imageFormat));
                }
                else
                {
                    var alt = element.GetAttribute("alt");
                    if (!string.IsNullOrEmpty(alt))
                    {
                        // ReSharper disable once RedundantSuppressNullableWarningExpression
                        segments.Add(new(alt!, format));
                    }
                }

                return;
            }
            case "svg":
            {
                var imageFormat = format;
                imageFormat.Image = ParseSvgElement(element);
                segments.Add(new("\uFFFC", imageFormat));
                return;
            }
        }

        var isBlock = IsBlockElement(tag);
        if (isBlock)
        {
            EnsureNewline(segments, format);
        }

        switch (tag)
        {
            case "li":
            {
                var parent = element.ParentElement?.LocalName;
                var depth = GetListDepth(element);
                var indent = depth > 0 ? new(' ', depth * 2) : "";
                var bulletFormat = newFormat;
                bulletFormat.ListDepth = depth;
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

                    segments.Add(new($"{indent}{index}. ", bulletFormat));
                }
                else
                {
                    var bullet = depth switch
                    {
                        0 => "●",
                        1 => "○",
                        _ => "■"
                    };
                    segments.Add(new($"{indent}{bullet} ", bulletFormat));
                }

                ProcessNode(element, newFormat, segments, inPre, settings);
                EnsureNewline(segments, format);
                return;
            }
            case "a":
            {
                ProcessNode(element, newFormat, segments, inPre, settings);
                var href = element.GetAttribute("href");
                if (!string.IsNullOrEmpty(href))
                {
                    if (!element.TextContent.AsSpan().Trim().Equals(href.AsSpan(), StringComparison.Ordinal))
                    {
                        segments.Add(new($" ({href})", format));
                    }
                }

                return;
            }
            case "q":
            {
                segments.Add(new("\u201C", format));
                ProcessNode(element, newFormat, segments, inPre, settings);
                segments.Add(new("\u201D", format));
                return;
            }
            case "td" or "th":
            {
                ProcessNode(element, newFormat, segments, inPre, settings);
                var nextSibling = element.NextElementSibling;
                if (nextSibling is { LocalName: "td" or "th" })
                {
                    segments.Add(new("\t", format));
                }

                return;
            }
            case "tr":
            {
                ProcessNode(element, newFormat, segments, inPre, settings);
                EnsureNewline(segments, format);
                return;
            }
            case "col":
                return;
            case "pre":
                ProcessNode(element, newFormat, segments, true, settings);
                break;
            default:
                ProcessNode(element, newFormat, segments, inPre, settings);
                break;
        }

        if (isBlock)
        {
            EnsureNewline(segments, format);
        }
    }

    internal static void ApplyElementFormatting(IElement element, string tag, ref FormatState format) =>
        ApplyElementFormatting(element, tag, ref format, out _);

    internal static bool ApplyWhiteSpace(Dictionary<string, string>? declarations, ref FormatState format, bool inPre)
    {
        if (declarations == null ||
            !declarations.TryGetValue("white-space", out var value))
        {
            return inPre;
        }

        var ws = value.AsSpan().Trim();
        if (ws.Equals("pre", StringComparison.OrdinalIgnoreCase) ||
            ws.Equals("pre-wrap", StringComparison.OrdinalIgnoreCase) ||
            ws.Equals("break-spaces", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (ws.Equals("nowrap", StringComparison.OrdinalIgnoreCase))
        {
            format.NoWrap = true;
        }

        return inPre;
    }

    internal static void ApplyElementFormatting(IElement element, string tag, ref FormatState format, out Dictionary<string, string>? declarations)
    {
        switch (tag)
        {
            case "b" or "strong" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "th" or "caption" or "dt":
                format.Bold = true;
                break;
            case "i" or "em" or "cite" or "dfn" or "var":
                format.Italic = true;
                break;
            case "u" or "ins":
                format.UnderlineStyle ??= UnderlineValues.Single;
                break;
            case "s" or "strike" or "del":
                format.Strikethrough = true;
                break;
            case "a":
                format.UnderlineStyle ??= UnderlineValues.Single;
                format.Color ??= "0563C1";
                format.LinkUrl = element.GetAttribute("href");
                format.LinkTitle = element.GetAttribute("title");
                break;
            case "sup":
                format.Superscript = true;
                format.Subscript = false;
                break;
            case "sub":
                format.Subscript = true;
                format.Superscript = false;
                break;
            case "mark":
                format.Color ??= "000000";
                format.BackgroundColor ??= "FFFF00";
                break;
            case "code" or "kbd" or "samp":
                format.FontFamily ??= "Courier New";
                break;
            case "small":
                format.FontSizePt = Math.Round((format.FontSizePt ?? 12) * 0.8, 2);
                break;
        }

        if (tag == "font")
        {
            var fontColor = element.GetAttribute("color");
            if (fontColor != null)
            {
                var parsed = ColorParser.Parse(fontColor);
                if (parsed != null)
                {
                    format.Color = parsed;
                }
            }

            var fontFace = element.GetAttribute("face");
            if (fontFace != null)
            {
                format.FontFamily = fontFace;
            }

            var fontSize = element.GetAttribute("size");
            if (fontSize != null && double.TryParse(fontSize, CultureInfo.InvariantCulture, out var size))
            {
                format.FontSizePt = size;
            }
        }

        ApplyInlineStyle(element, ref format, out declarations);
    }

    static void ApplyInlineStyle(IElement element, ref FormatState format, out Dictionary<string, string>? declarations)
    {
        var style = element.GetAttribute("style");
        if (style == null)
        {
            declarations = null;
            return;
        }

        declarations = StyleParser.Parse(style);

        if (declarations.TryGetValue("font-weight", out var fontWeight))
        {
            if (fontWeight is "bold" or "bolder" or "700" or "800" or "900")
            {
                format.Bold = true;
            }
        }

        if (declarations.TryGetValue("font-style", out var fontStyle))
        {
            if (fontStyle is "italic" or "oblique")
            {
                format.Italic = true;
            }
        }

        if (declarations.TryGetValue("text-decoration", out var textDecoration))
        {
            var tdSpan = textDecoration.AsSpan();
            var i = 0;
            while (i < tdSpan.Length)
            {
                while (i < tdSpan.Length && tdSpan[i] == ' ')
                {
                    i++;
                }

                if (i >= tdSpan.Length)
                {
                    break;
                }

                var start = i;
                while (i < tdSpan.Length && tdSpan[i] != ' ')
                {
                    i++;
                }

                var token = tdSpan[start..i];
                if (token.Equals("underline", StringComparison.OrdinalIgnoreCase))
                {
                    format.UnderlineStyle ??= UnderlineValues.Single;
                }
                else if (token.Equals("line-through", StringComparison.OrdinalIgnoreCase))
                {
                    format.Strikethrough = true;
                }
            }
        }

        if (declarations.TryGetValue("text-decoration-style", out var decorationStyle) &&
            format.UnderlineStyle != null)
        {
            if (decorationStyle.Equals("dotted", StringComparison.OrdinalIgnoreCase))
            {
                format.UnderlineStyle = UnderlineValues.Dotted;
            }
            else if (decorationStyle.Equals("dashed", StringComparison.OrdinalIgnoreCase))
            {
                format.UnderlineStyle = UnderlineValues.Dash;
            }
            else if (decorationStyle.Equals("wavy", StringComparison.OrdinalIgnoreCase))
            {
                format.UnderlineStyle = UnderlineValues.Wave;
            }
            else if (decorationStyle.Equals("double", StringComparison.OrdinalIgnoreCase))
            {
                format.UnderlineStyle = UnderlineValues.Double;
            }
        }

        if (declarations.TryGetValue("color", out var color))
        {
            var parsed = ColorParser.Parse(color);
            if (parsed != null)
            {
                format.Color = parsed;
            }
        }

        if (declarations.TryGetValue("font-size", out var fontSizeStr))
        {
            var parsed = StyleParser.ParseFontSize(fontSizeStr);
            if (parsed != null)
            {
                format.FontSizePt = parsed;
            }
        }

        if (declarations.TryGetValue("font-family", out var fontFamily))
        {
            format.FontFamily = fontFamily.Trim('\'', '"');
        }

        if (declarations.TryGetValue("background-color", out var bgColor) ||
            declarations.TryGetValue("background", out bgColor))
        {
            var parsed = ColorParser.Parse(bgColor);
            if (parsed != null)
            {
                format.BackgroundColor = parsed;
            }
        }

        if (declarations.TryGetValue("font-variant", out var fontVariant))
        {
            if (fontVariant.Contains("small-caps", StringComparison.OrdinalIgnoreCase))
            {
                format.SmallCaps = true;
            }
        }

        if (declarations.TryGetValue("text-transform", out var textTransform))
        {
            if (textTransform.Equals("uppercase", StringComparison.OrdinalIgnoreCase))
            {
                format.TextTransform = "uppercase";
            }
            else if (textTransform.Equals("lowercase", StringComparison.OrdinalIgnoreCase))
            {
                format.TextTransform = "lowercase";
            }
            else if (textTransform.Equals("capitalize", StringComparison.OrdinalIgnoreCase))
            {
                format.TextTransform = "capitalize";
            }
        }

        if (declarations.TryGetValue("border", out var borderVal))
        {
            format.Border = StyleParser.ParseBorder(borderVal);
        }

        if (declarations.TryGetValue("text-shadow", out var textShadow) &&
            !textShadow.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            format.Shadow = true;
        }

        if (declarations.TryGetValue("letter-spacing", out var letterSpacing) &&
            !letterSpacing.Equals("normal", StringComparison.OrdinalIgnoreCase))
        {
            var twips = StyleParser.ParseLengthToTwips(letterSpacing);
            if (twips != null)
            {
                format.CharacterSpacingTwips = twips;
            }
        }

        if (declarations.TryGetValue("vertical-align", out var verticalAlign))
        {
            switch (verticalAlign)
            {
                case "super":
                    format.Superscript = true;
                    format.Subscript = false;
                    break;
                case "sub":
                    format.Subscript = true;
                    format.Superscript = false;
                    break;
            }
        }
    }

    internal static bool IsBlockElement(string tag) =>
        tag is "p" or "div" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6"
            or "ul" or "ol" or "li" or "blockquote" or "pre" or "table"
            or "tr" or "hr" or "section" or "article" or "header" or "footer"
            or "nav" or "aside" or "main" or "figure" or "figcaption" or "details"
            or "summary" or "address" or "dt" or "dd" or "dl"
            or "caption" or "tbody" or "thead" or "tfoot"
            or "body" or "html";

    internal static bool IsInterBlockWhitespace(IText textNode)
    {
        var parent = textNode.ParentElement;
        if (parent == null)
        {
            return false;
        }

        var hasChildren = false;
        foreach (var child in parent.Children)
        {
            hasChildren = true;
            if (!IsBlockElement(child.LocalName))
            {
                return false;
            }
        }

        return hasChildren;
    }

    internal static string CollapseWhitespace(string text)
    {
        // Fast scan: if every whitespace char is already a single ' ' with a non-space neighbor,
        // the input is already in collapsed form and can be returned without allocation.
        var lastWasSpace = false;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (char.IsWhiteSpace(c))
            {
                if (c != ' ' || lastWasSpace)
                {
                    return CollapseWhitespaceSlow(text);
                }

                lastWasSpace = true;
            }
            else
            {
                lastWasSpace = false;
            }
        }

        return text;
    }

    static string CollapseWhitespaceSlow(string text)
    {
        var builder = new StringBuilder(text.Length);
        var lastWasSpace = false;
        foreach (var c in text)
        {
            if (char.IsWhiteSpace(c))
            {
                if (!lastWasSpace)
                {
                    builder.Append(' ');
                    lastWasSpace = true;
                }
            }
            else
            {
                builder.Append(c);
                lastWasSpace = false;
            }
        }

        return builder.ToString();
    }

    static void EnsureNewline(List<TextSegment> segments, FormatState format)
    {
        if (segments.Count == 0)
        {
            return;
        }

        var last = segments[^1];
        if (!last.Text.EndsWith('\n'))
        {
            segments.Add(new("\n", format));
        }
    }

    internal static int GetListDepth(IElement listItem)
    {
        var depth = 0;
        var current = listItem.ParentElement;
        while (current != null)
        {
            if (current.LocalName is "ul" or "ol")
            {
                depth++;
            }

            current = current.ParentElement;
        }

        // Subtract 1 because the immediate parent ul/ol is depth 0
        return Math.Max(0, depth - 1);
    }

    static void TrimTrailingNewlines(List<TextSegment> segments)
    {
        while (segments.Count > 0 && segments[^1].Text == "\n")
        {
            segments.RemoveAt(segments.Count - 1);
        }
    }

    internal static ImageData? ParseImageSrc(IElement element)
    {
        var src = element.GetAttribute("src");
        if (src == null || !src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var commaIndex = src.IndexOf(',');
        if (commaIndex < 0)
        {
            return null;
        }

        var meta = src.AsSpan(5, commaIndex - 5);
        if (!meta.EndsWith(";base64".AsSpan(), StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var contentType = meta[..^7].ToString();
        var base64 = src.AsSpan(commaIndex + 1);

#if NETFRAMEWORK
        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(base64.ToString());
        }
        catch (FormatException)
        {
            return null;
        }
#else
        var bytes = new byte[base64.Length];
        if (!Convert.TryFromBase64Chars(base64, bytes, out var bytesWritten))
        {
            return null;
        }

        Array.Resize(ref bytes, bytesWritten);
#endif

        var (width, height) = ImageResolver.ParseImageDimensions(element);
        return new(bytes, contentType, width, height);
    }

    internal static ImageData ParseSvgElement(IElement element)
    {
        if (!element.HasAttribute("xmlns"))
        {
            element.SetAttribute("xmlns", "http://www.w3.org/2000/svg");
        }

        using var writer = new StringWriter();
        element.ToHtml(writer, XhtmlMarkupFormatter.Instance);
        var svgBytes = Encoding.UTF8.GetBytes(writer.ToString());

        var (width, height) = ResolveSvgDimensions(element);
        return new(svgBytes, "image/svg+xml", width, height);
    }

    static (int? Width, int? Height) ResolveSvgDimensions(IElement element)
    {
        var width = ParseSvgDimension(element.GetAttribute("width"));
        var height = ParseSvgDimension(element.GetAttribute("height"));
        if (width != null && height != null)
        {
            return (width, height);
        }

        var viewBox = element.GetAttribute("viewBox");
        if (viewBox == null)
        {
            return (width, height);
        }

        var dims = ParseViewBox(viewBox);
        return (width ?? dims.Width, height ?? dims.Height);
    }

    static int? ParseSvgDimension(string? value)
    {
        if (value == null)
        {
            return null;
        }

        var valueSpan = value.AsSpan().Trim();
        if (valueSpan.EndsWith("px".AsSpan(), StringComparison.OrdinalIgnoreCase))
        {
            valueSpan = valueSpan[..^2];
        }

        if (double.TryParse(valueSpan, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
        {
            return (int)Math.Round(result);
        }

        return null;
    }

    static (int? Width, int? Height) ParseViewBox(string viewBox)
    {
        var span = viewBox.AsSpan();
        Span<int> starts = stackalloc int[4];
        Span<int> ends = stackalloc int[4];
        var count = 0;
        var i = 0;
        while (i < span.Length && count < 4)
        {
            while (i < span.Length && (span[i] == ' ' || span[i] == ','))
            {
                i++;
            }

            if (i >= span.Length)
            {
                break;
            }

            var start = i;
            while (i < span.Length && span[i] != ' ' && span[i] != ',')
            {
                i++;
            }

            starts[count] = start;
            ends[count] = i;
            count++;
        }

        if (count >= 4 &&
            double.TryParse(span.Slice(starts[2], ends[2] - starts[2]), NumberStyles.Float, CultureInfo.InvariantCulture, out var w) &&
            double.TryParse(span.Slice(starts[3], ends[3] - starts[3]), NumberStyles.Float, CultureInfo.InvariantCulture, out var h))
        {
            return ((int)Math.Round(w), (int)Math.Round(h));
        }

        return (null, null);
    }
}
