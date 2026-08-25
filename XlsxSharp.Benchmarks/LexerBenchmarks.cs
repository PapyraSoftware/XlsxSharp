using System;
using BenchmarkDotNet.Attributes;

namespace XlsxSharp.Benchmarks;

// Tokenizes every formula of the Enron dataset with each lexer alone, no parsing. Attributes
// FormulaParserBenchmarks' XlsxSharp.Parser vs XlsxSharp.Parser (Pratt) gap to lexing vs parsing:
// the RDS side wins lexing with a generated DFA table (XlsxSharp.Parser.Rolex), while Pratt hand-
// rolls a branch-chain scanner (XlsxSharp.Parser.Pratt.Lexer) - if this benchmark shows most of the
// gap, the scanner is the thing to optimize; if it doesn't, look at the parser's hot path instead.
[MemoryDiagnoser]
public class LexerBenchmarks
{
    private string[] _formulas = [];

    [GlobalSetup]
    public void GlobalSetup()
    {
        _formulas = FormulaCorpus.LoadEnron();
    }

    [Benchmark(Baseline = true, Description = "Rolex (RDS) lexer")]
    public int RolexLexer()
    {
        int tokenCount = 0;
        foreach (string formula in this._formulas)
        {
            tokenCount += XlsxSharp.Parser.Rolex.RolexLexer.GetTokensA1(formula.AsSpan()).Count;
        }

        return tokenCount;
    }

    [Benchmark(Description = "Pratt lexer")]
    public int PrattLexer()
    {
        // Unlike RolexLexer.GetTokensA1 (which reports a malformed formula as a trailing error
        // token instead of throwing), the hand-rolled Pratt lexer throws directly on malformed
        // input (unterminated string, unpaired surrogate, ...) - so this needs the same per-formula
        // try/catch as the parser benchmarks in FormulaParserBenchmarks to keep going past those.
        XlsxSharp.Parser.Pratt.Lexer lexer = new();
        int tokenCount = 0;
        foreach (string formula in this._formulas)
        {
            try
            {
                lexer.Reset(formula);
                XlsxSharp.Parser.Pratt.Token token;
                do
                {
                    token = lexer.Consume();
                    tokenCount++;
                } while (token.Type != XlsxSharp.Parser.Pratt.TokenType.Eof);
            }
            catch (Exception) { }
        }

        return tokenCount;
    }
}
