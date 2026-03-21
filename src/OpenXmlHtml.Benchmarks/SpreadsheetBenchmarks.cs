using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlHtml;

#pragma warning disable CA1822 // Mark members as static - BenchmarkDotNet requires instance methods

namespace OpenXmlHtml.Benchmarks;

[MemoryDiagnoser]
public class SpreadsheetBenchmarks
{
    const string SimpleFormatting = "<b>Bold</b> and <i>italic</i> text";

    const string RichCell = """
        <b>Status:</b> <span style="color: green">Active</span><br>
        <i>Updated:</i> <code>2024-01-15</code><br>
        <a href="https://example.com">Details</a>
        """;

    [Benchmark]
    public InlineString ToInlineString_Simple() =>
        SpreadsheetHtmlConverter.ToInlineString(SimpleFormatting);

    [Benchmark]
    public InlineString ToInlineString_RichCell() =>
        SpreadsheetHtmlConverter.ToInlineString(RichCell);
}
