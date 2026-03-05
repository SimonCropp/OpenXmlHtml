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

**Pipeline**: HTML string → AngleSharp DOM → `List<TextSegment>` → OpenXml elements

- `HtmlSegmentParser` (internal) — Parses HTML via AngleSharp into flat `TextSegment(string Text, FormatState Format)` list. Handles block elements (newline separators), inline formatting, tables (tab-separated cells), lists (bullet/numbered prefixes), links (URL appended), whitespace collapsing, and `<pre>` preservation.
- `SpreadsheetHtmlConverter` (public) — Converts segments to `InlineString` with spreadsheet `Run` elements for xlsx cells.
- `WordHtmlConverter` (public) — Converts segments to `Paragraph`/`Run` elements for docx. Splits on `\n` segments to create paragraph boundaries. Also provides `ConvertToDocx` (string or stream) and `ConvertFileToDocx`.
- `ColorParser`, `StyleParser` (internal) — Parse CSS colors (hex/named/rgb) and inline style attributes using `ReadOnlySpan<char>` for minimal allocations.

## Key Conventions

- **Multi-target**: Library targets `net48;net10.0`. Tests target `net10.0;net48`. Uses [Polyfill](https://github.com/SimonCropp/Polyfill) + `System.Memory` for span support on net48.
- **Namespace conflicts**: `DocumentFormat.OpenXml.Spreadsheet` and `DocumentFormat.OpenXml.Wordprocessing` share type names (`Run`, `Bold`, `Color`, etc.). Resolved via global using aliases in `GlobalUsings.cs` (e.g., `SpreadsheetRun`, `SpreadsheetBold`).
- **Code style**: EditorConfig enforces `var` everywhere, expression bodies on methods/properties/constructors, braces always required, file-scoped namespaces, no access modifiers on types. `TreatWarningsAsErrors` is on.
- **Strong naming**: Uses `key.snk` via ProjectDefaults. `InternalsVisibleTo` in `AssemblyInfo.cs` requires full public key.
- **Central package management**: All versions in `src/Directory.Packages.props`.
- **LangVersion preview**: C# preview features are available (raw string literals, collection expressions, etc.).
