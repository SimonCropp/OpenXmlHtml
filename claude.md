# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Test

```bash
# Build
cd src && dotnet build OpenXmlHtml.slnx

# Run all tests (both net10.0 and net48)
cd src && dotnet test OpenXmlHtml.slnx

# Run single framework
cd src && dotnet test OpenXmlHtml.slnx --framework net10.0

# Run single test
cd src && dotnet test OpenXmlHtml.slnx --filter "TestName"
```

The solution file is at `src/OpenXmlHtml.slnx`. All commands must run from the `src/` directory.

Building the test project triggers MarkdownSnippets, which updates `readme.md` from `#region` snippets in `src/OpenXmlHtml.Tests/Samples/`.

## Verify Snapshot Testing

Tests use [Verify](https://github.com/VerifyTests/Verify) with NUnit. When test output changes:
- `.received.*` files appear next to `.verified.*` files
- Accept changes by copying received to verified: `cp file.received.xml file.verified.xml`
- Tests producing docx binaries or floating-point output use `.UniqueForTargetFrameworkAndVersion()` since output differs between net10.0 and net48
- Custom Verify converters in `VerifyOpenXmlConverter.cs` serialize OpenXml objects to `.OuterXml`

## Architecture

Two code paths exist for Word output:

**Flat segment path** (`ToParagraphs`): HTML → AngleSharp DOM → `List<TextSegment>` → `Paragraph`/`Run` elements. Simple, no `MainDocumentPart` required, but no tables, headings styles, numbering, or style mapping.

**DOM-based path** (`ToElements`/`AppendHtml`/`ConvertToDocx`): HTML → AngleSharp DOM → `WordContentBuilder.Build` → `List<OpenXmlElement>`. Full-featured: tables, heading styles, real list numbering, CSS class→style mapping, paragraph spacing, images, footnotes, bookmarks.

### Key internal classes

- `HtmlSegmentParser` — Parses HTML via AngleSharp into flat `TextSegment(string Text, FormatState Format)` list. Used by both `ToParagraphs` (flat path) and `SpreadsheetHtmlConverter`.
- `WordContentBuilder` — DOM-based Word builder. Walks the AngleSharp DOM tree via `ProcessElement`/`ProcessChildren`, accumulates runs in `WordBuildContext`, flushes paragraphs with full styling (heading styles, list numbering, CSS class styles, paragraph spacing).
- `ImageResolver` — Resolves `<img>` sources: `data:` URIs (always allowed), HTTP/HTTPS URLs (checked against `HtmlConvertSettings.WebImages` policy), local files (checked against `LocalImages` policy). Returns `ImageData` or null (alt text fallback).
- `WordNumberingBuilder` — Creates `NumberingDefinitionsPart` with bullet/decimal abstract numbering definitions. Each `<ul>`/`<ol>` gets its own numbering instance; nested lists increment `ilvl`.
- `WordStyleLookup` — Reads `StyleDefinitionsPart` to build a map from CSS class names to Word paragraph/character style IDs (case-insensitive).
- `ParagraphFormatState` — Holds paragraph-level CSS properties (margins, text-indent, line-height, text-align) parsed from inline `style` attributes, applied at paragraph flush time.

### Public API classes

- `WordHtmlConverter` (public) — All Word conversion entry points. `ToParagraphs` uses flat segments; everything else delegates to `WordContentBuilder`. All methods have overloads accepting `HtmlConvertSettings`.
- `SpreadsheetHtmlConverter` (public) — Converts segments to `InlineString` with spreadsheet `Run` elements for xlsx cells.
- `ImagePolicy` (public) — Controls which image sources are allowed: `Deny()`, `AllowAll()`, `SafeDomains(...)`, `SafeDirectories(...)`, `Filter(predicate)`.
- `HtmlConvertSettings` (public) — Settings for image resolution: `WebImages`/`LocalImages` policies, optional `HttpClient`.
- `ColorParser`, `StyleParser` (internal) — Parse CSS colors (hex/named/rgb), inline style attributes, and CSS lengths (pt/px/em/in/cm/mm → twips).

## Test Organization

Tests are organized by feature area. Each supported HTML element and CSS property should have a dedicated test. When adding a new feature, add a corresponding test in the appropriate file.

### Test file → feature mapping

| Test File | Covers |
|---|---|
| `WordBasicTests` | `b`, `strong`, `i`, `em`, `u`, `ins`, `s`, `strike`, `del`, `sub`, `sup`, `br`, HTML entities |
| `WordBlockTests` | `p`, `div`, `h1`–`h6`, `blockquote`, `pre`, `hr`, `ul`/`ol`/`li` (text prefix path), page breaks |
| `WordHeadingTests` | `h1`–`h6` heading styles |
| `WordColorAndFontTests` | `color`, `font-size`, `font-family`, `font` attributes, `small`, `code`/`kbd`/`samp`, named/hex/rgb colors |
| `WordMiscElementTests` | `abbr`, `acronym`, `time`, `q`, `figure`/`figcaption`, `svg`, `article`, `section`, `nav`, `main`, `header`, `footer`, `aside`, `dfn`, `details`/`summary`, `address`, `dl`/`dt`/`dd` |
| `WordTableTests` | `table`, `tr`, `td`, `th`, `colspan`, `rowspan`, `thead`/`tbody`/`tfoot`, `caption`, nested tables |
| `WordTableStyleTests` | Cell `padding`/`width`/`background-color`/`vertical-align`, table `width`/`background-color`/`padding`, `cellpadding`/`bgcolor`/`width` HTML attributes |
| `WordAnchorTests` | `a` (hyperlinks, internal `#id` links), `id` attribute bookmarks |
| `WordNestedTests` | Deeply nested formatting combinations |
| `WordEdgeCaseTests` | Whitespace collapsing, malformed HTML, unclosed tags, unknown tags, image alt fallback |
| `WordParagraphSpacingTests` | `margin`, `text-indent`, `text-align`, `line-height` |
| `WordBackgroundColorTests` | `background-color` on runs/paragraphs, `background` shorthand, `<mark>` element |
| `WordUnderlineTests` | `text-decoration-style` variants (dotted, dashed, wavy, double), `<u>`/`<ins>` tags, spreadsheet underline |
| `WordStyleMappingTests` | CSS `class` → Word paragraph/character style mapping |
| `WordListNumberingTests` | Real Word numbering (`NumberingDefinitionsPart`), nested lists, separate list restart, fallback |
| `WordRemoteImageTests` | `ImagePolicy` (Deny/AllowAll/SafeDomains/Filter/SafeDirectories), `FakeImageHandler` |
| `WordConvertToDocxTests` | Full docx output: images, SVG, footnotes, page breaks, CSS styles, lists, tables |
| `WordConvertFileTests` | `ConvertFileToDocx` file I/O |
| `WordIntegrationTests` | `AppendHtml`, `ToParagraphs` rich document scenarios |
| `WordStyleComboTests` | Single large docx exercising all features together |
| `ImagePolicyTests` | `ImagePolicy` unit tests (Deny/AllowAll/SafeDomains/SafeDirectories/Filter) |
| `StyleParserTests` | CSS parsing, `ParseFontSize`, `ParseLengthToTwips` |
| `ColorParserTests` | Hex/named/RGB color parsing |
| `Spreadsheet*Tests` | Mirror of Word tests for spreadsheet-supported features |

### Test requirements

**Every new feature and bug fix must have a dedicated test.** Do not rely on combo tests or incidental coverage. The test should be named after the specific feature or bug (e.g., `SmallCapsTag`, `BorderShorthand`, `NestedListNumberingRestart`).

1. Find the appropriate test file from the table above (or create a new `Word<Feature>Tests.cs`)
2. Add a test method named after the feature
3. For features requiring `MainDocumentPart` (styles, numbering, images), use `ConvertToDocx` or `AppendHtml` with a `MainDocumentPart`
4. For simple formatting, `ToParagraphs` is sufficient
5. Run test → copy `.received.*` to `.verified.*` → run again to confirm
6. Update the test file mapping table above if you create a new test file

## Key Conventions

- **Multi-target**: Library targets `net48;net10.0`. Tests target `net10.0;net48`. Uses [Polyfill](https://github.com/SimonCropp/Polyfill) + `System.Memory` for span support on net48.
- **Namespace conflicts**: `DocumentFormat.OpenXml.Spreadsheet` and `DocumentFormat.OpenXml.Wordprocessing` share type names (`Run`, `Bold`, `Color`, etc.). Resolved via global using aliases in `GlobalUsings.cs` (e.g., `SpreadsheetRun`, `SpreadsheetBold`).
- **Code style**: EditorConfig enforces `var` everywhere, expression bodies on methods/properties/constructors, braces always required, file-scoped namespaces, no access modifiers on types. `TreatWarningsAsErrors` is on.
- **Strong naming**: Uses `key.snk` via ProjectDefaults. `InternalsVisibleTo` in `AssemblyInfo.cs` requires full public key.
- **Central package management**: All versions in `src/Directory.Packages.props`.
- **LangVersion preview**: C# preview features are available (raw string literals, collection expressions, etc.).
