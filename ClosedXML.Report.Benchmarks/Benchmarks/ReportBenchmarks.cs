using System.Reflection;
using BenchmarkDotNet.Attributes;
using ClosedXML.Report.Benchmarks.Models;

namespace ClosedXML.Report.Benchmarks.Benchmarks;

[MemoryDiagnoser]
public class ReportBenchmarks
{
    private Customer _customer = null!;
    private byte[] _templateData = [];
    const string ResourceName = "ClosedXML.Report.Benchmarks.Resources.Benchmark.xlsx";

    [GlobalSetup]
    public void Setup()
    {
        // Load embedded resource outside of the benchmark
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
        if (stream == null)
            throw new Exception($"Resource {ResourceName} not found");

        // Create a memory stream to hold the resource data
        var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        _templateData = memoryStream.ToArray();
        memoryStream.Position = 0;

        // Prepare data
        var dataBuilder = new DataBuilder();
        _customer = dataBuilder.Create();
    }

    [Benchmark]
    public void ReportGeneration()
    {
        using var memoryStream = new MemoryStream(_templateData);
        using var xlTemplate = new XLTemplate(memoryStream);

        // Add variables and generate report
        xlTemplate.AddVariable(_customer);
        xlTemplate.Generate();

        //Not testing the save, as it has a high overhead.
        //xlTemplate.SaveAs(Path.Combine("Output", "BenchmarkOutput.xlsx"));
    }
}