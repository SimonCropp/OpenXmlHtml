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

        var span = value.AsSpan().Trim();

        if (span[0] == '#')
        {
            var hex = span[1..];
            if (!IsValidHex(hex))
            {
                return null;
            }

            if (hex.Length == 3)
            {
                return string.Concat(
                    hex[0].ToString(), hex[0].ToString(),
                    hex[1].ToString(), hex[1].ToString(),
                    hex[2].ToString(), hex[2].ToString()).ToUpperInvariant();
            }

            if (hex.Length == 6)
            {
                return hex.ToString().ToUpperInvariant();
            }

            if (hex.Length == 8)
            {
                // #RRGGBBAA — strip alpha, keep RGB
                return hex[..6].ToString().ToUpperInvariant();
            }

            return null;
        }

        if (span.StartsWith("rgb".AsSpan(), StringComparison.OrdinalIgnoreCase))
        {
            return ParseRgb(span);
        }

        return namedColors.GetValueOrDefault(span.ToString());
    }

    static string? ParseRgb(ReadOnlySpan<char> span)
    {
        var start = span.IndexOf('(');
        var end = span.IndexOf(')');
        if (start < 0 || end < 0)
        {
            return null;
        }

        var inner = span[(start + 1)..end];

        var firstComma = inner.IndexOf(',');
        if (firstComma < 0)
        {
            return null;
        }

        var rest = inner[(firstComma + 1)..];
        var secondComma = rest.IndexOf(',');
        if (secondComma < 0)
        {
            return null;
        }

        // For rgba(r,g,b,a), extract only r,g,b — ignore alpha
        var bSpan = rest[(secondComma + 1)..];
        var thirdComma = bSpan.IndexOf(',');
        if (thirdComma >= 0)
        {
            bSpan = bSpan[..thirdComma];
        }

        if (!int.TryParse(inner[..firstComma].Trim().ToString(), out var r) ||
            !int.TryParse(rest[..secondComma].Trim().ToString(), out var g) ||
            !int.TryParse(bSpan.Trim().ToString(), out var b))
        {
            return null;
        }

        return $"{Clamp(r):X2}{Clamp(g):X2}{Clamp(b):X2}";
    }

    static int Clamp(int value) =>
        Math.Max(0, Math.Min(255, value));

    static bool IsValidHex(ReadOnlySpan<char> hex)
    {
        if (hex.Length is not (3 or 6 or 8))
        {
            return false;
        }

        foreach (var c in hex)
        {
            if (!char.IsAsciiHexDigit(c))
            {
                return false;
            }
        }

        return true;
    }
}
