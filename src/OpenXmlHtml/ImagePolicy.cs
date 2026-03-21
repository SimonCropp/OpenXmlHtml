namespace OpenXmlHtml;

/// <summary>
/// Controls which image sources are allowed during HTML conversion.
/// </summary>
public sealed class ImagePolicy
{
    readonly ImagePolicyKind kind;
    readonly Func<string, bool>? filter;

    ImagePolicy(ImagePolicyKind kind, Func<string, bool>? filter = null)
    {
        this.kind = kind;
        this.filter = filter;
    }

    /// <summary>
    /// Rejects all remote/local images. This is the default policy.
    /// </summary>
    public static ImagePolicy Deny() =>
        new(ImagePolicyKind.Deny);

    /// <summary>
    /// Allows images from any source.
    /// </summary>
    public static ImagePolicy AllowAll() =>
        new(ImagePolicyKind.AllowAll);

    /// <summary>
    /// Allows local images only from the specified directories.
    /// </summary>
    public static ImagePolicy SafeDirectories(params string[] directories)
    {
        var normalized = directories
            .Select(NormalizeDirPath)
            .ToArray();
        return new(ImagePolicyKind.SafeList, source =>
        {
            var path = source;
            if (path.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    path = new Uri(path).LocalPath;
                }
                catch
                {
                    return false;
                }
            }

            try
            {
                var fullPath = Path.GetFullPath(path);
                return normalized.Any(dir => fullPath.StartsWith(dir, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        });
    }

    /// <summary>
    /// Allows web images only from the specified domains (exact or subdomain match).
    /// </summary>
    public static ImagePolicy SafeDomains(params string[] domains)
    {
        var domainList = domains.ToArray();
        return new(ImagePolicyKind.SafeList, source =>
        {
            if (!Uri.TryCreate(source, UriKind.Absolute, out var uri))
            {
                return false;
            }

            var host = uri.Host;
            return domainList.Any(d =>
                string.Equals(host, d, StringComparison.OrdinalIgnoreCase) ||
                host.EndsWith("." + d, StringComparison.OrdinalIgnoreCase));
        });
    }

    /// <summary>
    /// Allows images matching a custom predicate.
    /// </summary>
    public static ImagePolicy Filter(Func<string, bool> predicate) =>
        new(ImagePolicyKind.Filter, predicate);

    internal bool IsAllowed(string source) =>
        kind switch
        {
            ImagePolicyKind.Deny => false,
            ImagePolicyKind.AllowAll => true,
            ImagePolicyKind.SafeList or ImagePolicyKind.Filter => filter!(source),
            _ => false
        };

    static string NormalizeDirPath(string dir)
    {
        var fullPath = Path.GetFullPath(dir);
        if (!fullPath.EndsWith(Path.DirectorySeparatorChar) &&
            !fullPath.EndsWith(Path.AltDirectorySeparatorChar))
        {
            fullPath += Path.DirectorySeparatorChar;
        }

        return fullPath;
    }
}

enum ImagePolicyKind
{
    Deny,
    AllowAll,
    SafeList,
    Filter
}
