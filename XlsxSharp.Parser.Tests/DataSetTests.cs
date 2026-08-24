using System.Diagnostics;

namespace XlsxSharp.Parser.Tests;

public class DataSetTests
{
    [Test]
    public async Task EnronDataSetIsParseable()
    {
        await this.Assert_formulas_parsed_or_not_as_expected(
            "./data/enron/formulas.csv",
            ["./data/enron/invalid-external-cell-reference.csv", "./data/enron/known-fails.csv"]
        );
    }

    [Test]
    public async Task EusesDataSetIsParseable()
    {
        await this.Assert_formulas_parsed_or_not_as_expected(
            "./data/euses/formulas.csv",
            ["./data/euses/invalid-external-cell-reference.csv", "./data/euses/known-fails.csv"]
        );
    }

    [Test]
    public async Task ContributionsDataSetIsParseable()
    {
        await this.Assert_formulas_parsed_or_not_as_expected(
            "./data/contributions/formulas.csv",
            []
        );
    }

    private async Task Assert_formulas_parsed_or_not_as_expected(
        string input,
        string[] badFormulaPaths
    )
    {
        HashSet<string> badFormulas = [];
        foreach (string badFormulaPath in badFormulaPaths)
        {
            badFormulas.UnionWith(DataSets.ReadCsv(badFormulaPath));
        }

        // Read to memory before the parsing to measure only parsing.
        List<string> formulas = [.. DataSets.ReadCsv(input)];
        Stopwatch sw = Stopwatch.StartNew();
        int formulaCount = 0;
        foreach (string formula in formulas)
        {
            formulaCount++;
            try
            {
                _ = FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(
                    formula,
                    new Ctx(),
                    new F()
                );
                await Assert.That(badFormulas.Contains(formula)).IsFalse().Because(formula);
            }
            catch (Exception e)
            {
                await Assert
                    .That(badFormulas.Contains(formula))
                    .IsTrue()
                    .Because($"Parsing formula '{formula}' failed: {e.Message}");
            }
        }

        sw.Stop();
        double averageLength = formulas.Sum(x => x.Length) / (double)formulas.Count;
        TestContext.Current!.Output.WriteLine(
            $"Parsed {formulaCount} formulas (Average length {averageLength:F1}) in {sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds * 1000d / formulaCount:N3}μs/formula)."
        );
    }
}
