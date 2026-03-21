namespace OpenXmlHtml;

/// <summary>
/// Settings for controlling remote image resolution during HTML conversion.
/// </summary>
public class HtmlConvertSettings
{
    /// <summary>
    /// Policy for local file images (file:// URIs and filesystem paths). Default is Deny.
    /// </summary>
    public ImagePolicy LocalImages { get; init; } = ImagePolicy.Deny();

    /// <summary>
    /// Policy for web images (http:// and https:// URIs). Default is Deny.
    /// </summary>
    public ImagePolicy WebImages { get; init; } = ImagePolicy.Deny();

    /// <summary>
    /// Optional HttpClient for downloading web images. If null, a shared static instance is used.
    /// </summary>
    public HttpClient? HttpClient { get; init; }
}
