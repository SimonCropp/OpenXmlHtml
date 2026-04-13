static class StyleParser
{
    internal static Dictionary<string, string> Parse(string? style)
    {
        var result = new Dictionary<string, string>(8, StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(style))
        {
            return result;
        }

        var remaining = style.AsSpan();
        while (remaining.Length > 0)
        {
            var semiIndex = remaining.IndexOf(';');
            var declaration = semiIndex >= 0 ? remaining[..semiIndex] : remaining;
            remaining = semiIndex >= 0 ? remaining[(semiIndex + 1)..] : default;

            declaration = declaration.Trim();
            if (declaration.IsEmpty)
            {
                continue;
            }

            var colonIndex = declaration.IndexOf(':');
            if (colonIndex <= 0)
            {
                continue;
            }

            var property = declaration[..colonIndex].Trim();
            var value = declaration[(colonIndex + 1)..].Trim();
            if (!property.IsEmpty && !value.IsEmpty)
            {
                result[property.ToString()] = value.ToString();
            }
        }

        return result;
    }

    internal static double? ParseFontSize(string value)
    {
        var span = value.AsSpan().Trim();

        if (TryParseSuffix(span, "pt", out var pt))
        {
            return pt;
        }

        if (TryParseSuffix(span, "px", out var px))
        {
            return px * 0.75;
        }

        if (TryParseSuffix(span, "em", out var em))
        {
            return em * 12;
        }

        if (double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return raw;
        }

        if (span.Equals("xx-small", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }

        if (span.Equals("x-small", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }

        if (span.Equals("small", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }

        if (span.Equals("medium", StringComparison.OrdinalIgnoreCase))
        {
            return 12;
        }

        if (span.Equals("large", StringComparison.OrdinalIgnoreCase))
        {
            return 14;
        }

        if (span.Equals("x-large", StringComparison.OrdinalIgnoreCase))
        {
            return 18;
        }

        if (span.Equals("xx-large", StringComparison.OrdinalIgnoreCase))
        {
            return 24;
        }

        return null;
    }

    internal static int? ParseLengthToTwips(string value)
    {
        var span = value.AsSpan().Trim();

        // 1pt = 20 twips
        if (TryParseSuffix(span, "pt", out var pt))
        {
            return (int)Math.Round(pt * 20);
        }

        // 1px ≈ 0.75pt = 15 twips
        if (TryParseSuffix(span, "px", out var px))
        {
            return (int)Math.Round(px * 15);
        }

        // 1em = 12pt = 240 twips
        if (TryParseSuffix(span, "em", out var em))
        {
            return (int)Math.Round(em * 240);
        }

        // 1in = 1440 twips
        if (TryParseSuffix(span, "in", out var inches))
        {
            return (int)Math.Round(inches * 1440);
        }

        // 1cm = 567 twips
        if (TryParseSuffix(span, "cm", out var cm))
        {
            return (int)Math.Round(cm * 567);
        }

        // 1mm = 56.7 twips
        if (TryParseSuffix(span, "mm", out var mm))
        {
            return (int)Math.Round(mm * 56.7);
        }

        // Bare number treated as px
        if (double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return (int)Math.Round(raw * 15);
        }

        return null;
    }

    internal static int? ParseLengthToPixels(string value)
    {
        var twips = ParseLengthToTwips(value);
        return twips == null ? null : (int)Math.Round(twips.Value / 15d);
    }

    internal static (int? Top, int? Right, int? Bottom, int? Left) ParseMarginShorthand(string value)
    {
        var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length switch
        {
            1 => (ParseLengthToTwips(parts[0]), ParseLengthToTwips(parts[0]),
                  ParseLengthToTwips(parts[0]), ParseLengthToTwips(parts[0])),
            2 => (ParseLengthToTwips(parts[0]), ParseLengthToTwips(parts[1]),
                  ParseLengthToTwips(parts[0]), ParseLengthToTwips(parts[1])),
            3 => (ParseLengthToTwips(parts[0]), ParseLengthToTwips(parts[1]),
                  ParseLengthToTwips(parts[2]), ParseLengthToTwips(parts[1])),
            4 => (ParseLengthToTwips(parts[0]), ParseLengthToTwips(parts[1]),
                  ParseLengthToTwips(parts[2]), ParseLengthToTwips(parts[3])),
            _ => (null, null, null, null)
        };
    }

    internal static JustificationValues? ParseTextAlign(string value)
    {
        var span = value.AsSpan().Trim();
        if (span.Equals("left", StringComparison.OrdinalIgnoreCase))
        {
            return JustificationValues.Left;
        }

        if (span.Equals("center", StringComparison.OrdinalIgnoreCase))
        {
            return JustificationValues.Center;
        }

        if (span.Equals("right", StringComparison.OrdinalIgnoreCase))
        {
            return JustificationValues.Right;
        }

        if (span.Equals("justify", StringComparison.OrdinalIgnoreCase))
        {
            return JustificationValues.Both;
        }

        return null;
    }

    static readonly HashSet<string> borderStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        "solid", "dotted", "dashed", "double", "none", "hidden",
        "groove", "ridge", "inset", "outset"
    };

    internal static BorderInfo? ParseBorder(string value)
    {
        var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        int? widthEighths = null;
        BorderValues? style = null;
        string? color = null;

        foreach (var part in parts)
        {
            if (borderStyles.Contains(part))
            {
                style = ParseBorderStyle(part);
            }
            else if (widthEighths == null)
            {
                var w = ParseBorderWidth(part);
                if (w != null)
                {
                    widthEighths = w;
                    continue;
                }

                // Not a width, try as color
                var parsed = ColorParser.Parse(part);
                if (parsed != null)
                {
                    color = parsed;
                }
            }
            else
            {
                var parsed = ColorParser.Parse(part);
                if (parsed != null)
                {
                    color = parsed;
                }
            }
        }

        if (style == null && widthEighths == null && color == null)
        {
            return null;
        }

        if (style == BorderValues.None)
        {
            return new(0, BorderValues.None, null);
        }

        return new(widthEighths ?? 4, style ?? BorderValues.Single, color);
    }

    static BorderValues ParseBorderStyle(string part)
    {
        var span = part.AsSpan();
        if (span.Equals("solid", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.Single;
        }

        if (span.Equals("dotted", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.Dotted;
        }

        if (span.Equals("dashed", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.Dashed;
        }

        if (span.Equals("double", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.Double;
        }

        if (span.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            span.Equals("hidden", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.None;
        }

        if (span.Equals("groove", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.ThreeDEngrave;
        }

        if (span.Equals("ridge", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.ThreeDEmboss;
        }

        if (span.Equals("inset", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.Inset;
        }

        if (span.Equals("outset", StringComparison.OrdinalIgnoreCase))
        {
            return BorderValues.Outset;
        }

        return BorderValues.Single;
    }

    static int? ParseBorderWidth(string value)
    {
        var span = value.AsSpan().Trim();

        if (span.Equals("thin", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }

        if (span.Equals("medium", StringComparison.OrdinalIgnoreCase))
        {
            return 12;
        }

        if (span.Equals("thick", StringComparison.OrdinalIgnoreCase))
        {
            return 18;
        }

        // 1pt = 8 eighths
        if (TryParseSuffix(span, "pt", out var pt))
        {
            return Math.Max(1, (int)Math.Round(pt * 8));
        }

        // 1px ≈ 0.75pt = 6 eighths
        if (TryParseSuffix(span, "px", out var px))
        {
            return Math.Max(1, (int)Math.Round(px * 6));
        }

        if (double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return raw == 0 ? 0 : Math.Max(1, (int)Math.Round(raw * 6));
        }

        return null;
    }

    static bool TryParseSuffix(ReadOnlySpan<char> span, string suffix, out double result)
    {
        if (span.EndsWith(suffix.AsSpan(), StringComparison.OrdinalIgnoreCase))
        {
            return double.TryParse(
                span[..^suffix.Length].Trim(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result);
        }

        result = 0;
        return false;
    }
}
