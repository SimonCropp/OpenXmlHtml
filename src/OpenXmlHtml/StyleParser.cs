using System.Globalization;

namespace OpenXmlHtml;

static class StyleParser
{
    internal static Dictionary<string, string> Parse(string? style)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(style))
        {
            return result;
        }

        foreach (var declaration in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var colonIndex = declaration.IndexOf(':');
            if (colonIndex <= 0)
            {
                continue;
            }

            var property = declaration[..colonIndex].Trim();
            var value = declaration[(colonIndex + 1)..].Trim();
            if (property.Length > 0 && value.Length > 0)
            {
                result[property] = value;
            }
        }

        return result;
    }

    internal static double? ParseFontSize(string value)
    {
        value = value.Trim().ToLowerInvariant();

        if (value.EndsWith("pt") &&
            double.TryParse(value[..^2].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var pt))
        {
            return pt;
        }

        if (value.EndsWith("px") &&
            double.TryParse(value[..^2].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var px))
        {
            return px * 0.75;
        }

        if (value.EndsWith("em") &&
            double.TryParse(value[..^2].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var em))
        {
            return em * 12;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return raw;
        }

        return value switch
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
}
