using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;

namespace XlsxSharp.Benchmarks;

/// <summary>
/// Loads the real-world formula corpora already checked into XlsxSharp.Parser.Tests, shared by
/// every benchmark class that needs a realistic formula workload.
/// </summary>
internal static class FormulaCorpus
{
    public static string[] LoadEnron() => [.. Read("data/enron/formulas.csv")];

    private static IEnumerable<string> Read(string filename)
    {
        CsvConfiguration config = new(CultureInfo.InvariantCulture) { HasHeaderRecord = false };
        using StreamReader reader = new(filename);
        using CsvReader csv = new(reader, config);
        foreach (FormulaRecord record in csv.GetRecords<FormulaRecord>())
        {
            yield return record.Text;
        }
    }

    private sealed record FormulaRecord([Index(0)] string Text);
}
