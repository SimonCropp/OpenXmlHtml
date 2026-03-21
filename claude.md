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

## Key Conventions

- **Multi-target**: Library targets `net48;net10.0`. Tests target `net10.0;net48`. Uses [Polyfill](https://github.com/SimonCropp/Polyfill) + `System.Memory` for span support on net48.
- **Namespace conflicts**: `DocumentFormat.OpenXml.Spreadsheet` and `DocumentFormat.OpenXml.Wordprocessing` share type names (`Run`, `Bold`, `Color`, etc.). Resolved via global using aliases in `GlobalUsings.cs` (e.g., `SpreadsheetRun`, `SpreadsheetBold`).
- **Code style**: EditorConfig enforces `var` everywhere, expression bodies on methods/properties/constructors, braces always required, file-scoped namespaces, no access modifiers on types. `TreatWarningsAsErrors` is on.
- **Strong naming**: Uses `key.snk` via ProjectDefaults. `InternalsVisibleTo` in `AssemblyInfo.cs` requires full public key.
- **Central package management**: All versions in `src/Directory.Packages.props`.
- **LangVersion preview**: C# preview features are available (raw string literals, collection expressions, etc.).
