using System;
using BenchmarkDotNet.Attributes;

namespace XlsxSharp.Benchmarks;

// Parses every formula of the Enron dataset (XlsxSharp.Parser.Tests/data/enron/formulas.csv) - the
// same real-world corpus ClosedXML.Parser's own README quotes a parsing throughput for - through
// both parsers. XlsxSharpAstFactory and ClosedXmlAstFactory build the identical BenchNode shape, so
// the numbers below isolate the parsers themselves, not factory overhead.
//
// Two contenders:
//  - XlsxSharp.Parser (Pratt): the project's own (and, since the recursive-descent parser this
//    replaced was removed, only) formula parser, in XlsxSharp.Parser.Pratt.
//  - ClosedXML.Parser: the real upstream NuGet package the removed recursive-descent parser was
//    originally forked from - the more meaningful baseline now that the fork itself is gone.
// Pratt's ParserFactory.Create builds the parselet table once, so it's created outside the loop and
// reused across formulas via ParseFormula - the way it's meant to be used - rather than paying that
// setup cost per formula like the ClosedXML.Parser benchmark incidentally does inside CellFormulaA1.
[MemoryDiagnoser]
public class FormulaParserBenchmarks
{
    private string[] _formulas = [];

    [GlobalSetup]
    public void GlobalSetup()
    {
        _formulas = FormulaCorpus.LoadEnron();
    }

    [Benchmark(Description = "XlsxSharp.Parser (Pratt)")]
    public int XlsxSharpPrattParser()
    {
        XlsxSharpAstFactory factory = new();
        XlsxSharp.Parser.Pratt.Parser<BenchNode, object?> parser =
            XlsxSharp.Parser.Pratt.ParserFactory.Create(factory);
        int parsed = 0;
        foreach (string formula in this._formulas)
        {
            try
            {
                parser.ParseFormula(formula, null);
                parsed++;
            }
            catch (Exception)
            {
                // The dataset contains formulas neither parser is expected to accept (e.g. dangling
                // #REF!). Throughput of successful parses is what's being measured here, not
                // coverage - that's XlsxSharp.Parser.Tests' job.
            }
        }

        return parsed;
    }

    [Benchmark(Baseline = true, Description = "ClosedXML.Parser")]
    public int ClosedXmlParser()
    {
        ClosedXmlAstFactory factory = new();
        int parsed = 0;
        foreach (string formula in this._formulas)
        {
            try
            {
                ClosedXML.Parser.FormulaParser<BenchNode, BenchNode, object?>.CellFormulaA1(
                    formula,
                    null,
                    factory
                );
                parsed++;
            }
            catch (Exception) { }
        }

        return parsed;
    }
}
