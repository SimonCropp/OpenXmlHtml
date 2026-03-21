static class StyleParser
{
    internal static Dictionary<string, string> Parse(string? style)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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

        if (double.TryParse(span.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return raw;
        }

        return span.ToString().ToLowerInvariant() switch
        {
            "xx-small" => 7,
            "x-small" => 8,
            "small" => 10,
            "medium" => 12,
            "large" => 14,
            "x-large" => 18,
            "xx-large" => 24,
            _ => null
        };
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
        if (double.TryParse(span.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return (int)Math.Round(raw * 15);
        }

        return null;
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

    internal static JustificationValues? ParseTextAlign(string value) =>
        value.Trim().ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "justify" => JustificationValues.Both,
            _ => null
        };

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
                style = part.ToLowerInvariant() switch
                {
                    "solid" => BorderValues.Single,
                    "dotted" => BorderValues.Dotted,
                    "dashed" => BorderValues.Dashed,
                    "double" => BorderValues.Double,
                    "none" or "hidden" => BorderValues.None,
                    "groove" => BorderValues.ThreeDEngrave,
                    "ridge" => BorderValues.ThreeDEmboss,
                    "inset" => BorderValues.Inset,
                    "outset" => BorderValues.Outset,
                    _ => BorderValues.Single
                };
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

    static int? ParseBorderWidth(string value)
    {
        var lower = value.Trim().ToLowerInvariant();
        switch (lower)
        {
            case "thin":
                return 4;
            case "medium":
                return 12;
            case "thick":
                return 18;
        }

        var span = lower.AsSpan();

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

        if (double.TryParse(lower, NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
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
                span[..^suffix.Length].Trim().ToString(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result);
        }

        result = 0;
        return false;
    }
}
