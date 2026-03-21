```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.7058/22H2/2022Update)
Intel Core i9-9880H CPU 2.30GHz, 1 CPU, 16 logical and 8 physical cores
.NET SDK 10.0.200
  [Host]     : .NET 10.0.4 (10.0.426.12010), X64 RyuJIT AVX2
  DefaultJob : .NET 10.0.4 (10.0.426.12010), X64 RyuJIT AVX2


```
| Method                       | Mean         | Error      | StdDev       | Gen0     | Gen1     | Allocated  |
|----------------------------- |-------------:|-----------:|-------------:|---------:|---------:|-----------:|
| ToParagraphs_SimpleInline    |     35.49 μs |   0.400 μs |     0.355 μs |   2.1973 |        - |   18.57 KB |
| ToParagraphs_RichParagraphs  |    101.55 μs |   1.031 μs |     0.964 μs |   5.1270 |   0.2441 |   43.23 KB |
| ToElements_SimpleInline      |     35.30 μs |   0.319 μs |     0.298 μs |   2.1362 |   0.0610 |   17.92 KB |
| ToElements_RichParagraphs    |    118.52 μs |   1.794 μs |     1.591 μs |   4.8828 |        - |   47.77 KB |
| ToElements_NestedList        |    107.05 μs |   0.721 μs |     0.602 μs |   5.1270 |   0.2441 |   42.15 KB |
| ToElements_Table             |    165.39 μs |   1.660 μs |     1.471 μs |   6.8359 |        - |   56.23 KB |
| ToElements_NestedTable       |     77.73 μs |   1.125 μs |     1.052 μs |   4.1504 |   0.2441 |    34.3 KB |
| ConvertToDocx_RichParagraphs |    605.09 μs |   1.770 μs |     1.655 μs |  11.7188 |        - |  120.54 KB |
| ConvertToDocx_Table          |    564.70 μs |  10.541 μs |    18.183 μs |  11.7188 |        - |  121.12 KB |
| ConvertToDocx_NestedList     |  1,192.03 μs |  20.208 μs |    16.874 μs |  31.2500 |   7.8125 |  288.45 KB |
| ConvertToDocx_LargeDocument  | 26,223.70 μs | 458.169 μs |   579.437 μs | 571.4286 | 428.5714 | 5968.05 KB |
| AppendHtml_RichParagraphs    |    601.20 μs |   8.991 μs |    10.354 μs |  11.7188 |        - |  120.54 KB |
| AppendHtml_LargeDocument     | 25,477.86 μs | 495.970 μs | 1,415.030 μs | 600.0000 | 400.0000 | 5967.59 KB |
