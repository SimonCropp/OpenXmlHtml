static class ColorParser
{
    static readonly Dictionary<string, string> namedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000",
        ["white"] = "FFFFFF",
        ["red"] = "FF0000",
        ["green"] = "008000",
        ["blue"] = "0000FF",
        ["yellow"] = "FFFF00",
        ["orange"] = "FFA500",
        ["purple"] = "800080",
        ["pink"] = "FFC0CB",
        ["gray"] = "808080",
        ["grey"] = "808080",
        ["brown"] = "A52A2A",
        ["cyan"] = "00FFFF",
        ["magenta"] = "FF00FF",
        ["lime"] = "00FF00",
        ["navy"] = "000080",
        ["teal"] = "008080",
        ["maroon"] = "800000",
        ["olive"] = "808000",
        ["aqua"] = "00FFFF",
        ["silver"] = "C0C0C0",
        ["fuchsia"] = "FF00FF",
        ["darkred"] = "8B0000",
        ["darkgreen"] = "006400",
        ["darkblue"] = "00008B",
        ["darkcyan"] = "008B8B",
        ["darkmagenta"] = "8B008B",
        ["darkorange"] = "FF8C00",
        ["darkviolet"] = "9400D3",
        ["lightblue"] = "ADD8E6",
        ["lightgreen"] = "90EE90",
        ["lightgray"] = "D3D3D3",
        ["lightgrey"] = "D3D3D3",
        ["lightyellow"] = "FFFFE0",
        ["lightpink"] = "FFB6C1",
        ["indianred"] = "CD5C5C",
        ["coral"] = "FF7F50",
        ["tomato"] = "FF6347",
        ["gold"] = "FFD700",
        ["khaki"] = "F0E68C",
        ["plum"] = "DDA0DD",
        ["violet"] = "EE82EE",
        ["orchid"] = "DA70D6",
        ["salmon"] = "FA8072",
        ["sienna"] = "A0522D",
        ["tan"] = "D2B48C",
        ["wheat"] = "F5DEB3",
        ["crimson"] = "DC143C",
        ["indigo"] = "4B0082",
        ["steelblue"] = "4682B4",
        ["skyblue"] = "87CEEB",
        ["slategray"] = "708090",
        ["slategrey"] = "708090"
    };

    internal static string? Parse(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        // ReSharper disable once RedundantSuppressNullableWarningExpression
        value = value!.Trim();

        if (value.StartsWith('#'))
        {
            var hex = value[1..];
            if (!IsValidHex(hex))
            {
                return null;
            }

            if (hex.Length == 3)
            {
                hex = string.Concat(hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]);
            }

            if (hex.Length is 6 or 8)
            {
                return hex.ToUpperInvariant();
            }

            return null;
        }

        if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
        {
            return ParseRgb(value);
        }

        return namedColors.GetValueOrDefault(value);
    }

    static string? ParseRgb(string value)
    {
        var start = value.IndexOf('(');
        var end = value.IndexOf(')');
        if (start < 0 || end < 0)
        {
            return null;
        }

        var parts = value[(start + 1)..end].Split(',');
        if (parts.Length < 3)
        {
            return null;
        }

        if (!int.TryParse(parts[0].Trim(), out var r) ||
            !int.TryParse(parts[1].Trim(), out var g) ||
            !int.TryParse(parts[2].Trim(), out var b))
        {
            return null;
        }

        return $"{Clamp(r):X2}{Clamp(g):X2}{Clamp(b):X2}";
    }

    static int Clamp(int value) =>
        Math.Max(0, Math.Min(255, value));

    static bool IsValidHex(string hex) =>
        hex.Length is 3 or 6 or 8 &&
        hex.All(char.IsAsciiHexDigit);
}
