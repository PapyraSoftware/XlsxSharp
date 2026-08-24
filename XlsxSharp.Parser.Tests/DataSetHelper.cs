using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using JetBrains.Annotations;

namespace XlsxSharp.Parser.Tests;

internal static class DataSets
{
    public static IEnumerable<string> ReadCsv(string filename)
    {
        CsvConfiguration config = new(CultureInfo.InvariantCulture) { HasHeaderRecord = false };
        using StreamReader reader = new(filename);
        using CsvReader csv = new(reader, config);
        IEnumerable<Formula> formulas = csv.GetRecords<Formula>();
        foreach (Formula formula in formulas)
        {
            yield return formula.Text;
        }
    }

    [UsedImplicitly]
    private record Formula([Index(0)] string Text);
}
