using BenchmarkDotNet.Running;

BenchmarkSwitcher.FromAssembly(typeof(OpenXmlHtml.Benchmarks.WordBenchmarks).Assembly).Run(args);
