using Antlr4.Runtime;
using Antlr4.Runtime.Atn;

namespace XlsxSharp.Parser.Tests;

internal static class AssertFormula
{
    /// <summary>
    /// Assert that a formula is parsed into a single childless node.
    /// </summary>
    public static async Task SingleNodeParsed<TNode>(string formula, TNode expectedNode)
        where TNode : AstNode
    {
        TNode node = (TNode)
            FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), new F());
        await Assert.That(node).IsEqualTo(expectedNode);
    }

    public static async Task CheckParsingErrorContains(string formula, string errorSubstring)
    {
        ParsingException ex = Assert.ThrowsExactly<ParsingException>(() =>
            FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), new F())
        );
        await Assert
            .That(ex.Message.Contains(errorSubstring))
            .IsTrue()
            .Because($"Error message '{ex.Message}' doesn't contain '{errorSubstring}'.");
    }

    /// <summary>
    /// Assert that text is recognized as a single token of a token type.
    /// </summary>
    /// <param name="tokenText">Text that should contain a single token.</param>
    /// <param name="tokenType">Expected token type, from <see cref="FormulaLexer"/> const .</param>
    public static async Task AssertTokenType(string tokenText, int tokenType)
    {
        CommonTokenStream commonTokenStream = new(
            new FormulaLexer(new AntlrInputStream(tokenText))
        );
        commonTokenStream.Fill();
        await Assert.That(commonTokenStream.Size).IsEqualTo(2);
        await Assert.That(commonTokenStream.Get(0).Type).IsEqualTo(tokenType);
        await Assert.That(commonTokenStream.Get(1).Type).IsEqualTo(FormulaLexer.Eof);
    }

    public static async Task CstParsed(string formula)
    {
        AntlrInputStream inputStream = new(formula);
        FormulaLexer lexer = new(inputStream);
        LexerErrorListener listener = new();
        lexer.RemoveErrorListeners();
        lexer.AddErrorListener(listener);
        CommonTokenStream commonTokenStream = new(lexer);
        FormulaParser parser = new(commonTokenStream, TextWriter.Null, TextWriter.Null)
        {
            Interpreter = { PredictionMode = PredictionMode.SLL },
        };
        parser.ErrorListeners.Clear();

        FormulaParser.FormulaContext res = parser.formula();

        await Assert
            .That(listener.ErrorStartIndex is null)
            .IsTrue()
            .Because($"{formula}  {listener.ErrorStartIndex}");
        await Assert.That(res.exception is null).IsTrue().Because($"{formula} {res.exception}");
    }

    /// <summary>
    /// Get tokens from ANTLR lexer. If there is an error in the <paramref name="formula"/>, insert error token.
    /// </summary>
    public static IReadOnlyList<Token> GetAntlrTokens(string formula)
    {
        FormulaLexer lexer = new(
            new CodePointCharStream(formula),
            TextWriter.Null,
            TextWriter.Null
        );
        LexerErrorListener listener = new();
        lexer.RemoveErrorListeners();
        lexer.AddErrorListener(listener);
        List<Token> tokens =
        [
            .. lexer
                .GetAllTokens()
                .Select(x => new Token(x.Type, x.StartIndex, x.StopIndex - x.StartIndex + 1)),
        ];
        if (listener.ErrorStartIndex is not null)
        {
            // Lexer tries to recover. That is good in most cases, but in our case, it's not very
            // compatible with Rolex lexer. Remove the tokens after recovery.
            int errorStartIndex = listener.ErrorStartIndex.Value;
            List<Token> tokensWithError =
            [
                .. tokens.Where(t => t.StartIndex < errorStartIndex),
                new(Token.ErrorSymbolId, errorStartIndex, 0),
            ];
            return tokensWithError;
        }

        tokens.Add(Token.EofSymbol(formula.Length));
        return tokens;
    }

    private class LexerErrorListener : IAntlrErrorListener<int>
    {
        internal int? ErrorStartIndex { get; private set; }

        public void SyntaxError(
            TextWriter output,
            IRecognizer recognizer,
            int offendingSymbol,
            int line,
            int charPositionInLine,
            string msg,
            RecognitionException e
        )
        {
            // Params don't provide access to the stream char index property directly, so pass it through
            this.ErrorStartIndex ??= ((Lexer)recognizer).TokenStartCharIndex;
        }
    }
}
