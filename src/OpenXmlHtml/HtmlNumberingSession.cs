namespace OpenXmlHtml;

/// <summary>
/// Mutable state shared across multiple <c>WordHtmlConverter.ToElements</c> calls within a
/// single logical render. Passing the same instance to every call ensures that bullet lists
/// across separate HTML fragments reuse one shared abstract numbering definition rather than each
/// allocating their own.
/// </summary>
public class HtmlNumberingSession
{
    /// <summary>
    /// The abstract numbering ID used for bullet lists, populated after the first call that
    /// encounters a <c>&lt;ul&gt;</c>. Can be pre-set to reuse an existing abstract definition.
    /// </summary>
    public int? BulletAbstractNumId { get; set; }
}
