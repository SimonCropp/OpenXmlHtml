namespace OpenXmlHtml;

/// <summary>
/// Settings for HTML-to-Word conversion: image policies, heading offset, and optional shared
/// numbering state.
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

    /// <summary>
    /// Offset applied to HTML heading levels (h1..h6) when mapping to Word Heading styles.
    /// For example, an offset of 1 maps &lt;h1&gt; to Heading2, &lt;h2&gt; to Heading3, etc. The result
    /// is clamped to [1, 9] so out-of-range inputs still produce a valid heading style. Default 0.
    /// </summary>
    public int HeadingLevelOffset { get; init; }

    /// <summary>
    /// Optional shared numbering session. When provided, bullet abstract numbering definitions
    /// are reused across multiple <c>WordHtmlConverter.ToElements</c> calls that share
    /// the same session, so N HTML fragments with bullet lists produce one abstract definition
    /// instead of N.
    /// </summary>
    public HtmlNumberingSession? NumberingSession { get; init; }
}
