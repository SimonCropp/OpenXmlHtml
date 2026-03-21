```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.7058/22H2/2022Update)
Intel Core i9-9880H CPU 2.30GHz, 1 CPU, 16 logical and 8 physical cores
.NET SDK 10.0.200
  [Host]     : .NET 10.0.4 (10.0.426.12010), X64 RyuJIT AVX2
  DefaultJob : .NET 10.0.4 (10.0.426.12010), X64 RyuJIT AVX2


```
| Method                  | Mean     | Error    | StdDev   | Gen0   | Gen1   | Allocated |
|------------------------ |---------:|---------:|---------:|-------:|-------:|----------:|
| ToInlineString_Simple   | 25.61 μs | 0.381 μs | 0.357 μs | 1.9226 | 0.0610 |  15.89 KB |
| ToInlineString_RichCell | 58.87 μs | 1.156 μs | 1.994 μs | 2.9297 |      - |   26.9 KB |
