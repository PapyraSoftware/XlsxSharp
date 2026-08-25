using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

public class LexerTests
{
    // ( [0-9] )+
    [Arguments("0")]
    [Arguments("1")]
    [Arguments("90")]
    [Arguments("00050")]
    // ( [0-9] )+ '.' ( [0-9]] )+
    [Arguments("0.0")]
    [Arguments("1.2")]
    [Arguments("0010.0020")]
    [Arguments("999.99")]
    // '.' ( [0-9]] )+
    [Arguments(".0")]
    [Arguments(".1")]
    [Arguments(".0001")]
    [Arguments(".987")]
    // ( [0-9] )+ [Ee] ( [0-9]] )+
    [Arguments("0e0")]
    [Arguments("0E0")]
    [Arguments("1e2")]
    [Arguments("1E2")]
    [Arguments("987e12")]
    // ( [0-9] )+ '.' ( [0-9]] )+ [Ee] ( [0-9]] )+
    [Arguments("0.0e4")]
    [Arguments("12.724e13")]
    [Arguments("12.3E2")]
    // '.' ( [0-9]] )+ [Ee] ( [0-9]] )+
    [Arguments(".0e0")]
    [Arguments(".1e2")]
    [Arguments(".987e54")]
    // ( [0-9] )+ [Ee] [+-] ( [0-9]] )+
    [Arguments("1e+7")]
    [Arguments("74e-32")]
    [Arguments("15E-0")]
    [Arguments("0e+0")]
    [Arguments("01e+7")]
    // ( [0-9] )+ '.' ( [0-9]] )+ [Ee] [+-] ( [0-9]] )+
    [Arguments("0.0e+0")]
    [Arguments("1.2e+3")]
    [Arguments("01.2e+3")]
    [Arguments("1.2E+3")]
    [Arguments("12.34e+56")]
    // '.' ( [0-9]] )+ [Ee] [+-] ( [0-9]] )+
    [Arguments(".0e+0")]
    [Arguments(".1e+2")]
    [Arguments(".12E+34")]
    [Arguments(".012e+034")]
    [Test]
    public async Task NumberOk(string input)
    {
        await AssertToken(TokenType.Number, input);
    }

    [Arguments("0e+")]
    [Arguments(".0e+")]
    [Test]
    public async Task NumberFails(string input)
    {
        await AssertFail(input, "Number");
    }

    [Arguments("\"\"")]
    [Arguments("\"Some text\"")]
    [Arguments("\"Some \"\" text\"")]
    [Arguments("\"\uD83E\uDD8A\"")] // Fox face through surrogates
    [Test]
    public async Task TextOk(string input)
    {
        await AssertToken(TokenType.Text, input);
    }

    [Arguments("\"")]
    [Arguments("\"text")]
    [Arguments("\"text\"\"")]
    [Arguments("\"Some \"\" text")]
    [Test]
    public async Task TextMustBeTerminated(string input)
    {
        await AssertFail(input, "unterminated literal");
    }

    [Arguments("\"\u0015\"")]
    [Test]
    public async Task TextMustBeContainXml10Characters(string input)
    {
        await AssertFail(input, "Invalid text character");
    }

    [Arguments("#DIV/0!")]
    [Arguments("#GETTING_DATA")]
    [Arguments("#N/A")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#REF!")]
    [Arguments("#VALUE!")]
    [Arguments("#ref!")]
    [Test]
    public async Task ErrorOk(string input)
    {
        await AssertToken(TokenType.Error, input);
    }

    [Test]
    public async Task LexerThrowsOnUnpairedSurrogates()
    {
        // Either Visual Studio or NUnit is converting invalid surrogates to -1/65536. O
        string[] invalidCodeUnits =
        [
            "\uD83E", // Unpaired high surrogate for Fox Face
            "\uD83E*", // Unpaired high surrogate for Fox Face
            "\uDD8A", // Unpaired low surrogate for Fox Face
            "\uDD8A*\"", // Unpaired low surrogate for Fox Face
            "\uDD8A\uD83E", // Low surrogate first
            "name\uD83E)", // Unpaired high surrogate after an ASCII identifier run
        ];
        foreach (string invalidText in invalidCodeUnits)
        {
            await AssertFail(invalidText, "surrogate");
        }
    }

    [Arguments("''")]
    [Arguments("'[1]Something'")]
    [Arguments("'Jane''s'")]
    [Arguments("'New York'")]
    [Arguments("'January 1st:December 31st'")]
    [Arguments("'[7]Year 20:Year 25'")]
    [Arguments("'[Book.xlsx]Year 20:Year 25'")]
    [Arguments("'[End*Near.xlsx]Final'")]
    [Arguments("''''''")]
    [Test]
    public async Task QIdentOk(string input)
    {
        await AssertToken(TokenType.QIdent, input);
    }

    [Arguments("'")]
    [Arguments("'Jane''s")]
    [Arguments("'''''")]
    [Test]
    public async Task QIdentMustBeTerminated(string input)
    {
        await AssertFail(input, "unterminated literal");
    }

    [Arguments("ABC")]
    [Arguments("A1")]
    [Arguments("$A$1")]
    [Arguments("AEF$A$1")]
    [Arguments("name")]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("true")]
    [Arguments("false")]
    [Arguments("?name")]
    [Arguments("\\name")]
    [Arguments("_name")]
    [Arguments("name?")]
    [Arguments("name\\")]
    [Arguments("name_")]
    [Arguments("some.name")]
    [Arguments("_xlfn.ACOT")]
    [Arguments("\u05D0\u05D1\u05E0")] // stone in hebrew - Letters from other languages
    [Arguments("\u05E9\u05B0\u05DC\u05D5\u05DD")] // shalom - A mark from other languages
    [Arguments("name\uD83E\uDD8A")] // ASCII run followed by a valid surrogate pair (fox face)
    [Test]
    public async Task IdentOk(string input)
    {
        await AssertToken(TokenType.Ident, input);
    }

    [Test]
    public async Task IdentStopsAtOperators()
    {
        Dictionary<TokenType, string> operators = new()
        {
            { TokenType.Bang, "!" },
            { TokenType.Comma, "," },
            { TokenType.Semicolon, ";" },
            { TokenType.Pow, "^" },
            { TokenType.Mul, "*" },
            { TokenType.Div, "/" },
            { TokenType.Plus, "+" },
            { TokenType.Minus, "-" },
            { TokenType.Concat, "&" },
            { TokenType.Equal, "=" },
            { TokenType.NotEqual, "<>" },
            { TokenType.Less, "<" },
            { TokenType.LessEqual, "<=" },
            { TokenType.Greater, ">" },
            { TokenType.GreaterEqual, ">=" },
            { TokenType.Percent, "%" },
            { TokenType.Range, ":" },
            { TokenType.Spill, "#" },
            { TokenType.Intersection, "@" },
            { TokenType.LeftParen, "(" },
            { TokenType.RightParen, ")" },
            { TokenType.LeftCurly, "{" },
            { TokenType.RightCurly, "}" },
            { TokenType.Whitespace, " " },
        };

        foreach ((TokenType opType, string opText) in operators)
        {
            string input = "name" + opText;
            Lexer lexer = new(input);

            Pratt.Token identToken = lexer.Consume();
            await Assert.That(identToken.Type).IsEqualTo(TokenType.Ident);
            await Assert.That(identToken.GetText(input).ToString()).IsEqualTo("name");

            Pratt.Token opToken = lexer.Consume();
            await Assert.That(opToken.Type).IsEqualTo(opType);
            await Assert.That(opToken.GetText(input).ToString()).IsEqualTo(opText);
        }
    }

    [Arguments("[1]")]
    [Arguments("[]")]
    [Arguments("['[]")]
    [Arguments("[Book1.xlsx]")]
    [Arguments("[#Data]")]
    [Arguments("[[#Data]]")]
    [Arguments("[[#Data],[#Headers]]")]
    [Arguments("['#]")]
    [Arguments("[985]")]
    [Arguments("[Jan:Dec]")]
    [Arguments("['['['[]")]
    [Arguments("[']']']]")]
    [Test]
    public async Task SquareIdentOk(string input)
    {
        await AssertToken(TokenType.SquareIdent, input);
    }

    [Arguments("[Ja[[a]]]")]
    [Test]
    public async Task SquareIdentAtMostTwoNestedBrackets(string input)
    {
        // Mostly to keep within something DFA can do.
        await AssertFail(input, "at most two nested square brackets");
    }

    [Arguments("[")]
    [Arguments("[text")]
    [Arguments("[[")]
    [Arguments("[a[b")]
    [Arguments("[Start[]and end")]
    [Arguments("[Start[']and end']")]
    [Test]
    public async Task SquareIdentMustBePaired(string input)
    {
        await AssertFail(input, "Unable to find closing square bracket");
    }

    [Arguments((int)TokenType.Bang, "!")]
    [Arguments((int)TokenType.Range, ":")]
    [Arguments((int)TokenType.Comma, ",")]
    [Arguments((int)TokenType.Semicolon, ";")]
    [Arguments((int)TokenType.Pow, "^")]
    [Arguments((int)TokenType.Mul, "*")]
    [Arguments((int)TokenType.Div, "/")]
    [Arguments((int)TokenType.Plus, "+")]
    [Arguments((int)TokenType.Minus, "-")]
    [Arguments((int)TokenType.Concat, "&")]
    [Arguments((int)TokenType.Percent, "%")]
    [Arguments((int)TokenType.Spill, "#")]
    [Arguments((int)TokenType.Intersection, "@")]
    [Arguments((int)TokenType.LeftParen, "(")]
    [Arguments((int)TokenType.RightParen, ")")]
    [Arguments((int)TokenType.LeftCurly, "{")]
    [Arguments((int)TokenType.RightCurly, "}")]
    [Test]
    public async Task SingleCharTokensOk(int token, string input)
    {
        // TODO: Dump xUnit. Can't even use internal classes as test fixtures, so I have to pass enum as int.
        await AssertToken((TokenType)token, input);
    }

    [Arguments((int)TokenType.Equal, "=")]
    [Arguments((int)TokenType.NotEqual, "<>")]
    [Arguments((int)TokenType.Less, "<")]
    [Arguments((int)TokenType.LessEqual, "<=")]
    [Arguments((int)TokenType.Greater, ">")]
    [Arguments((int)TokenType.GreaterEqual, ">=")]
    [Test]
    public async Task ComparisonTokensOk(int token, string input)
    {
        await AssertToken((TokenType)token, input);
    }

    [Arguments("\t")]
    [Arguments("\n")]
    [Arguments("\r")]
    [Arguments(" ")]
    [Arguments("\t \r\n")]
    [Test]
    public async Task WhitespaceOk(string input)
    {
        await AssertToken(TokenType.Whitespace, input);
    }

    private static async Task AssertToken(TokenType type, string input)
    {
        Lexer lexer = new(input);
        Pratt.Token token = lexer.Consume();
        await Assert.That(token.Type).IsEqualTo(type);
        await Assert.That(token.GetText(input).ToString()).IsEqualTo(input);
    }

    private static async Task AssertFail(string input, string exceptionSubstring)
    {
        Lexer lexer = new(input);
        ParsingException exception = Assert.ThrowsExactly<ParsingException>(() => lexer.Consume());
        Assert.NotNull(exception);
        await Assert
            .That(exception.Message.Contains(exceptionSubstring))
            .IsTrue()
            .Because(
                $"Expected to find '{exceptionSubstring}', but not found in '{exception.Message}'."
            );
    }
}
