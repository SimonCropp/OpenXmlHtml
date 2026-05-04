static class ImageResolver
{
    static readonly HttpClient sharedClient = new();

    static readonly Dictionary<string, string> extensionToContentType = new(StringComparer.OrdinalIgnoreCase)
    {
        [".jpg"] = "image/jpeg",
        [".jpeg"] = "image/jpeg",
        [".png"] = "image/png",
        [".gif"] = "image/gif",
        [".bmp"] = "image/bmp",
        [".tiff"] = "image/tiff",
        [".tif"] = "image/tiff",
        [".svg"] = "image/svg+xml",
        [".webp"] = "image/webp",
        [".ico"] = "image/x-icon"
    };

    internal static ImageData? Resolve(IElement element, HtmlConvertSettings? settings)
    {
        var src = element.GetAttribute("src");
        if (src == null)
        {
            return null;
        }

        if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            return HtmlSegmentParser.ParseImageSrc(element);
        }

        if (settings == null)
        {
            return null;
        }

        byte[]? bytes;
        string? contentType;

        if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            if (!settings.WebImages.IsAllowed(src))
            {
                return null;
            }

            var client = settings.HttpClient ?? sharedClient;
            try
            {
                using var response = client.GetAsync(src).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    return null;
                }

                bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                contentType = response.Content.Headers.ContentType?.MediaType;
            }
            catch
            {
                return null;
            }
        }
        else
        {
            if (!settings.LocalImages.IsAllowed(src))
            {
                return null;
            }

            var path = src;
            if (src.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                if (!Uri.TryCreate(src, UriKind.Absolute, out var uri))
                {
                    return null;
                }

                path = uri.LocalPath;
            }

            try
            {
                path = Path.GetFullPath(path);
                if (!File.Exists(path))
                {
                    return null;
                }

                bytes = File.ReadAllBytes(path);
                contentType = GuessContentType(path);
            }
            catch
            {
                return null;
            }
        }

        contentType ??= "image/png";

        var (width, height) = ParseImageDimensions(element);
        return new(bytes, contentType, width, height);
    }

    internal static (int? Width, int? Height) ParseImageDimensions(IElement element)
    {
        int? width = null;
        int? height = null;

        var style = element.GetAttribute("style");
        if (style != null)
        {
            var declarations = StyleParser.Parse(style);
            if (declarations.TryGetValue("width", out var cssWidth))
            {
                width = StyleParser.ParseLengthToPixels(cssWidth);
            }

            if (declarations.TryGetValue("height", out var cssHeight))
            {
                height = StyleParser.ParseLengthToPixels(cssHeight);
            }
        }

        if (width == null)
        {
            var widthAttr = element.GetAttribute("width");
            if (widthAttr != null && int.TryParse(widthAttr, out var w))
            {
                width = w;
            }
        }

        if (height == null)
        {
            var heightAttr = element.GetAttribute("height");
            if (heightAttr != null && int.TryParse(heightAttr, out var h))
            {
                height = h;
            }
        }

        return (width, height);
    }

    static string GuessContentType(string path) =>
        extensionToContentType.GetValueOrDefault(Path.GetExtension(path), "image/png");
}
