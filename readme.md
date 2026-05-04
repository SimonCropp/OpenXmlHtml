# <img src="/src/icon.png" height="30px"> OpenXmlHtml

[![Build status](https://img.shields.io/appveyor/build/SimonCropp/OpenXmlHtml)](https://ci.appveyor.com/project/SimonCropp/OpenXmlHtml)
[![NuGet Status](https://img.shields.io/nuget/v/OpenXmlHtml.svg?label=OpenXmlHtml)](https://www.nuget.org/packages/OpenXmlHtml/)

Converts HTML to OpenXml elements for use in xlsx and docx files.

Uses [AngleSharp](https://github.com/AangleSharp/AngleSharp) for HTML parsing and [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) for OpenXml generation.


## NuGet package

https://nuget.org/packages/OpenXmlHtml/


## Spreadsheet (xlsx)


### SetCellHtml

Set the value of a spreadsheet cell from HTML:

<!-- snippet: SetCellHtml -->
<a id='snippet-SetCellHtml'></a>
```cs
var cell = new SpreadsheetCell();
SpreadsheetHtmlConverter.SetCellHtml(cell, "<b>Hello</b> <i>World</i>");
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/SpreadsheetSamples.cs#L7-L12' title='Snippet source file'>snippet source</a> | <a href='#snippet-SetCellHtml' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ToInlineString

Get an `InlineString` for use in a cell:

<!-- snippet: ToInlineString -->
<a id='snippet-ToInlineString'></a>
```cs
var inlineString = SpreadsheetHtmlConverter.ToInlineString(
    "<b>Revenue:</b> <font color=\"#008000\">$1.2M</font>");
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/SpreadsheetSamples.cs#L20-L25' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToInlineString' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Lists in Cells

<!-- snippet: SpreadsheetList -->
<a id='snippet-SpreadsheetList'></a>
```cs
var inlineString = SpreadsheetHtmlConverter.ToInlineString(
    """
    <ul>
      <li><span style="color: green">Passed</span>: 47</li>
      <li><span style="color: red">Failed</span>: 3</li>
    </ul>
    """);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/SpreadsheetSamples.cs#L33-L43' title='Snippet source file'>snippet source</a> | <a href='#snippet-SpreadsheetList' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Rich Content

<!-- snippet: SpreadsheetRichContent -->
<a id='snippet-SpreadsheetRichContent'></a>
```cs
var cell = new SpreadsheetCell();
SpreadsheetHtmlConverter.SetCellHtml(cell,
    """
    <h2>Q1 Report</h2>
    <p>Revenue: <b style="color: green">$1.2M</b></p>
    <p>See <a href="https://example.com/report">full report</a></p>
    <table>
      <tr><th>Region</th><th>Sales</th></tr>
      <tr><td>North</td><td>$500K</td></tr>
      <tr><td>South</td><td>$700K</td></tr>
    </table>
    """);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/SpreadsheetSamples.cs#L51-L66' title='Snippet source file'>snippet source</a> | <a href='#snippet-SpreadsheetRichContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## Word (docx)


### ToParagraphs

Convert HTML to a list of `Paragraph` elements:

<!-- snippet: ToParagraphs -->
<a id='snippet-ToParagraphs'></a>
```cs
var paragraphs = WordHtmlConverter.ToParagraphs(
    """
    <h1>Report Title</h1>
    <p>This is a <b>bold</b> statement with <i>emphasis</i>.</p>
    """);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L7-L15' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToParagraphs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### AppendHtml

Append HTML content directly to a Word document body:

<!-- snippet: AppendHtml -->
<a id='snippet-AppendHtml'></a>
```cs
using var stream = new MemoryStream();
using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
var main = document.AddMainDocumentPart();
main.Document = new(new Body());

WordHtmlConverter.AppendHtml(
    main.Document.Body!,
    """
    <h1>Meeting Notes</h1>
    <p><i>Date: January 15, 2024</i></p>
    <ol>
      <li>Review <code>PR #123</code></li>
      <li>Update <u>documentation</u></li>
    </ol>
    """);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L23-L41' title='Snippet source file'>snippet source</a> | <a href='#snippet-AppendHtml' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Rich Document

<!-- snippet: WordRichDocument -->
<a id='snippet-WordRichDocument'></a>
```cs
var paragraphs = WordHtmlConverter.ToParagraphs(
    """
    <h2>Status Report</h2>
    <p>All systems <span style="color: green"><b>operational</b></span>.</p>
    <ul>
      <li>Server: <span style="color: green">OK</span></li>
      <li>Cache: <span style="color: red">Down</span></li>
    </ul>
    <p>Contact <a href="mailto:ops@example.com">ops team</a> for details.</p>
    """);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L49-L62' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordRichDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ConvertToDocx

Convert an HTML string directly to a docx file:

<!-- snippet: ConvertToDocx -->
<a id='snippet-ConvertToDocx'></a>
```cs
using var stream = new MemoryStream();
WordHtmlConverter.ConvertToDocx(
    """
    <h1>Report</h1>
    <p>This is a <b>bold</b> statement.</p>
    <ul>
      <li>Item one</li>
      <li>Item two</li>
    </ul>
    """,
    stream);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L70-L84' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConvertToDocx' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ConvertStreamToDocx

Convert an HTML stream to a docx stream:

<!-- snippet: ConvertStreamToDocx -->
<a id='snippet-ConvertStreamToDocx'></a>
```cs
using var htmlStream = new MemoryStream(
    "<h1>Report</h1><p>Content</p>"u8.ToArray());
using var docxStream = new MemoryStream();
WordHtmlConverter.ConvertToDocx(htmlStream, docxStream);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L93-L100' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConvertStreamToDocx' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ConvertFileToDocx

Convert an HTML file to a docx file:

<!-- snippet: ConvertFileToDocx -->
<a id='snippet-ConvertFileToDocx'></a>
```cs
WordHtmlConverter.ConvertFileToDocx(htmlPath, docxPath);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L292-L296' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConvertFileToDocx' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Headers and Footers

Set document headers and footers from HTML:

<!-- snippet: HeadersAndFooters -->
<a id='snippet-HeadersAndFooters'></a>
```cs
using var headerStream = new MemoryStream();
using var headerDoc = WordprocessingDocument.Create(
    headerStream, WordprocessingDocumentType.Document);
var headerMainPart = headerDoc.AddMainDocumentPart();
headerMainPart.Document = new(new Body());

WordHtmlConverter.AppendHtml(
    headerMainPart.Document.Body!,
    "<p>Document content</p>",
    headerMainPart);

WordHtmlConverter.SetHeader(headerMainPart,
    """<p style="text-align: center"><b>Company Name</b></p>""");

WordHtmlConverter.SetFooter(headerMainPart,
    """<p style="text-align: center; font-size: 9pt; color: gray">Confidential</p>""");
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L224-L243' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeadersAndFooters' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Headers and footers support all the same HTML elements and CSS properties as the document body (tables, formatting, images, etc.). Overloads accepting `HtmlConvertSettings` are also available.


### Remote Images

By default, only base64 data URI images are embedded. External images (HTTP/HTTPS URLs, local files) require explicit opt-in via `HtmlConvertSettings` and `ImagePolicy`:

<!-- snippet: RemoteImageSettings -->
<a id='snippet-RemoteImageSettings'></a>
```cs
// Configure which image sources are allowed
var settings = new HtmlConvertSettings
{
    WebImages = ImagePolicy.SafeDomains("cdn.example.com"),
    LocalImages = ImagePolicy.SafeDirectories(@"C:\Reports\Images")
};

// Pass settings to any conversion method
using var settingsStream = new MemoryStream();
WordHtmlConverter.ConvertToDocx(
    """
    <h1>Report</h1>
    <p><img src="https://cdn.example.com/logo.png" alt="Logo"></p>
    """,
    settingsStream,
    settings);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L166-L185' title='Snippet source file'>snippet source</a> | <a href='#snippet-RemoteImageSettings' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Available policies:

 * `ImagePolicy.Deny()` — Default. Rejects all remote/local images.
 * `ImagePolicy.AllowAll()` — Allows any source.
 * `ImagePolicy.SafeDomains("example.com", ...)` — Allows web images from specified domains (exact or subdomain match).
 * `ImagePolicy.SafeDirectories(@"C:\Images", ...)` — Allows local images from specified directories (path traversal protected).
 * `ImagePolicy.Filter(src => ...)` — Custom predicate.

All conversion methods (`ToParagraphs`, `ToElements`, `AppendHtml`, `ConvertToDocx`, `ConvertFileToDocx`) and their spreadsheet equivalents accept an `HtmlConvertSettings` parameter. A custom `HttpClient` can be provided via `HtmlConvertSettings.HttpClient`.


### Shared Numbering Session

When calling `ToElements` multiple times to build a single Word document, each call normally creates its own bullet abstract numbering definition. Passing the same `HtmlNumberingSession` via `HtmlConvertSettings.NumberingSession` makes all calls share one definition:

<!-- snippet: SharedNumberingSession -->
<a id='snippet-SharedNumberingSession'></a>
```cs
var session = new HtmlNumberingSession();
var settings = new HtmlConvertSettings
{
    NumberingSession = session
};

// Both fragments reuse one bullet abstract — one definition, two numbering instances.
var first = WordHtmlConverter.ToElements("<ul><li>a</li><li>b</li></ul>", mainPart, settings);
var second = WordHtmlConverter.ToElements("<ul><li>c</li><li>d</li></ul>", mainPart, settings);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L199-L211' title='Snippet source file'>snippet source</a> | <a href='#snippet-SharedNumberingSession' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Without a session each `ToElements` call allocates its own abstract, which is harmless but produces redundant definitions. A session is only needed when you call `ToElements` several times against the same `MainDocumentPart` within one render.


### CSS Class to Word Style Mapping

When converting with a `MainDocumentPart` that has a `StyleDefinitionsPart`, CSS class names are matched against Word style definitions:

<!-- snippet: StyleMapping -->
<a id='snippet-StyleMapping'></a>
```cs
// CSS class names are matched against Word styles in the document.
// Paragraph styles apply via ParagraphStyleId,
// character styles apply via RunStyle.
WordHtmlConverter.AppendHtml(
    styleBody,
    """
    <p class="Quote">This uses the Quote paragraph style</p>
    <p>Normal text with <span class="Emphasis">emphasized</span> word</p>
    """,
    styleMainPart);
```
<sup><a href='/src/OpenXmlHtml.Tests/Samples/WordSamples.cs#L266-L279' title='Snippet source file'>snippet source</a> | <a href='#snippet-StyleMapping' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

 * Paragraph styles (`w:type="paragraph"`) are applied via `ParagraphStyleId`
 * Character styles (`w:type="character"`) are applied via `RunStyle`
 * Style lookup is case-insensitive
 * Heading styles (`Heading1`–`Heading6`) take precedence over CSS class styles


## Supported HTML Elements

### Text Formatting

 * `b`, `strong` - Bold
 * `i`, `em`, `cite`, `dfn`, `var` - Italic
 * `u`, `ins` - Underline
 * `s`, `strike`, `del` - Strikethrough
 * `sub` - Subscript
 * `sup` - Superscript
 * `small` - Smaller font size
 * `code`, `kbd`, `samp` - Monospace font (Courier New)
 * `mark` - Highlight (yellow background shading)
 * `a` - Hyperlink (underline + blue; Word: `#id` links become bookmarks, external URLs become clickable Word hyperlinks including images inside links)


### Block Elements

 * `p`, `div` - Paragraphs / divisions (Word: `id` attribute creates bookmark)
 * `h1` - `h6` - Headings (Word: Heading1-Heading6 paragraph styles; Spreadsheet: bold)
 * `blockquote` - Block quotation (Word: `cite` attribute creates footnote)
 * `pre` - Preformatted text (whitespace preserved)
 * `hr` - Horizontal rule
 * `article`, `aside`, `section`, `header`, `footer`, `nav`, `main` - Semantic blocks


### Lists

 * `ul`, `ol`, `li` - Unordered (bullet) and ordered (numbered) lists
   * Word (via `ToElements`/`ConvertToDocx`): real `NumberingDefinitionsPart` with `ListParagraph` style — proper bullets and numbering rendered by Word, supporting nested lists and separate numbering sequences
   * Word (via `ToParagraphs`): text prefix fallback (●/○/■ for bullets, 1./2./3. for ordered)
   * Spreadsheet: text prefix bullets and numbers
 * `type` attribute on `<ol>`: `1` (decimal), `a` (lower-alpha), `A` (upper-alpha), `i` (lower-roman), `I` (upper-roman) (Word)
 * `start` attribute on `<ol>`: starting number (e.g., `<ol start="5">`) (Word)
 * `list-style-type` CSS: `decimal`, `lower-alpha`/`lower-latin`, `upper-alpha`/`upper-latin`, `lower-roman`, `upper-roman` (Word)
 * `reversed` attribute on `<ol>`: countdown numbering (Word, via text prefix fallback)


### Tables

 * `table`, `tr`, `td`, `th` - Table structure
   * Word: real `Table`/`TableRow`/`TableCell` elements with borders
   * Spreadsheet: tab-separated cells, newline-separated rows
 * `colspan`, `rowspan` - Cell spanning (Word)
 * `thead`, `tbody`, `tfoot` - Table sections
 * `caption` - Table caption (bold)
 * `cellpadding` - HTML attribute for default cell padding (Word)
 * `bgcolor` - HTML attribute for cell background color (Word)
 * `width` - HTML attribute for cell width (Word)
 * Nested tables supported (Word)
 * Cell CSS: `padding`, `width`, `background-color`, `vertical-align` (top/middle/bottom) (Word)
 * Table CSS: `width`, `background-color`, `padding` (default cell padding) (Word)
 * Row CSS: `height` or HTML `height` attribute on `<tr>` (Word)


### Inline / Other

 * `br` - Line break
 * `span` - Generic inline container
 * `font` - Font styling (color, size, face attributes)
 * `time` - Time element
 * `abbr`, `acronym` - Abbreviations (Word: `title` attribute creates footnote)
 * `q` - Inline quotation (smart quotes)
 * `img` - Image (base64 data URIs always embedded; HTTP/local images via `HtmlConvertSettings`; alt text fallback)
 * `figure`, `figcaption` - Figure and caption
 * `svg` - Inline SVG (embedded as image in Word with PNG fallback)
 * `dl`, `dt`, `dd` - Definition list (dt is bold)


### CSS Style Attribute

Inline `style` attributes are supported:

 * `font-weight: bold` (or 700-900)
 * `font-style: italic`
 * `text-decoration: underline`, `text-decoration: line-through`
 * `text-decoration-style` - Underline variant: `dotted`, `dashed`, `wavy`, `double` (Word)
 * `border` - Border on runs, paragraphs, table cells, and tables (Word)
   * Shorthand: `border: 1px solid black`
   * Per-side: `border-top`, `border-right`, `border-bottom`, `border-left`
   * Styles: `solid`, `dotted`, `dashed`, `double`, `groove`, `ridge`, `inset`, `outset`, `none`
   * Table: `border: none` or HTML `border="0"` removes default table borders
 * `color` - Text color (hex, named, rgb())
 * `font-size` - Font size (pt, px, em, keywords)
 * `font-family` - Font family
 * `vertical-align: super`, `vertical-align: sub`
 * `page-break-before: always` - Page break before element (Word)
 * `page-break-after: always` - Page break after element (Word)
 * `text-align` - Paragraph alignment: left, center, right, justify (Word)
 * `margin` - Paragraph spacing and indentation (Word, supports shorthand and individual sides)
 * `margin-top`, `margin-bottom` - Paragraph spacing before/after (Word)
 * `margin-left`, `margin-right` - Paragraph indentation (Word)
 * `text-indent` - First line indent or hanging indent (Word)
 * `line-height` - Line spacing: numeric multiple (1.5), percentage (150%), or fixed length (18pt) (Word)
 * `background-color` - Background shading on runs and paragraphs (Word); also on table cells
 * `font-variant: small-caps` - Small capitals (Word)
 * `text-transform` - Transform text: `uppercase`, `lowercase`, `capitalize` (Word)
 * `writing-mode` - Text direction: `vertical-rl`, `vertical-lr` (Word, on paragraphs and table cells)
 * `direction: rtl` - Right-to-left text direction (Word)

CSS length units supported: `pt`, `px`, `em`, `in`, `cm`, `mm`.


### CSS Class Attribute

The `class` attribute maps CSS class names to Word styles defined in the document's `StyleDefinitionsPart`. See [CSS Class to Word Style Mapping](#css-class-to-word-style-mapping) above.


### Color Formats

 * Hex: `#RGB`, `#RRGGBB`, `#RRGGBBAA`
 * Named: all 148 CSS named colors (`red`, `blue`, `cornflowerblue`, `rebeccapurple`, etc.)
 * RGB: `rgb(255, 0, 0)`
 * RGBA: `rgba(255, 0, 0, 0.5)` (alpha channel is parsed but not applied — Word doesn't support color transparency)


## Why Not altChunk?

Word's altChunk (w:altChunk) mechanism can embed raw HTML inside a .docx file and have Word convert it at open time. This approach was considered and rejected for several reasons:

 * Deferred and lossy conversion — Word processes the HTML when the file is opened, replacing the altChunk with its own interpretation. The output varies by Word version and is not round-trip stable.
 * Undocumented tag support — Microsoft has never published a conformance spec for the HTML subset that Word's import filter accepts. Supported tags are determined empirically, and behaviour differs across versions.
 * Poor ecosystem compatibility — Many .docx consumers (LibreOffice, headless processors, other libraries) have partial or broken altChunk support. Files are not reliably portable.
 * No programmatic access — Because the HTML is opaque to the container, the content cannot be inspected, modified, or processed without first opening the file in Word.

OpenXmlHtml instead parses HTML with AngleSharp and emits native OOXML elements directly. This means the output is valid, stable, portable OOXML from the moment it is written — with no dependency on Word being installed, no conversion step, and predictable behaviour across all .docx consumers.


## Icon

https://thenounproject.com/icon/dog-6292156/
