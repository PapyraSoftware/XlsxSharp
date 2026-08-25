using System;
using BenchmarkDotNet.Attributes;

namespace XlsxSharp.Benchmarks;

// Parses every formula of the Enron dataset (XlsxSharp.Parser.Tests/data/enron/formulas.csv) - the
// same real-world corpus ClosedXML.Parser's own README quotes a parsing throughput for - through
// all three parsers. XlsxSharpAstFactory and ClosedXmlAstFactory build the identical BenchNode
// shape, so the numbers below isolate the parsers themselves, not factory overhead.
//
// Three contenders:
//  - XlsxSharp.Parser: our recursive-descent FormulaParser, a vendored fork of ClosedXML.Parser.
//  - XlsxSharp.Parser (Pratt): its in-progress replacement in XlsxSharp.Parser.Pratt, built to be
//    faster than the recursive-descent one it's meant to replace.
//  - ClosedXML.Parser: the real upstream NuGet package the fork started from.
// Pratt's ParserFactory.Create builds the parselet table once, so it's created outside the loop and
// reused across formulas via ParseFormula - the way it's meant to be used - rather than paying that
// setup cost per formula like the other two benchmarks incidentally do inside CellFormulaA1.
[MemoryDiagnoser]
public class FormulaParserBenchmarks
{
    private string[] _formulas = [];

    [GlobalSetup]
    public void GlobalSetup()
    {
        _formulas = FormulaCorpus.LoadEnron();
    }

    [Benchmark(Baseline = true, Description = "XlsxSharp.Parser")]
    public int XlsxSharpParser()
    {
        XlsxSharpAstFactory factory = new();
        int parsed = 0;
        foreach (string formula in this._formulas)
        {
            try
            {
                XlsxSharp.Parser.FormulaParser<BenchNode, BenchNode, object?>.CellFormulaA1(
                    formula,
                    null,
                    factory
                );
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
            catch (Exception) { }
        }

        return parsed;
    }

    [Benchmark(Description = "ClosedXML.Parser")]
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
