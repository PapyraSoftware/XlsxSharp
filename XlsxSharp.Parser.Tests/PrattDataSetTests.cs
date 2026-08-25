using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests;

/// <summary>
/// Regression coverage for the pratt parser (<see cref="XlsxSharp.Parser.Pratt"/>) - the only
/// formula parser in this project, now that <see cref="FormulaParser{TScalarValue,TNode,TContext}"/>
/// (the recursive-descent implementation this replaced) has been removed - against the real-world
/// formula corpora under <c>data/</c>.
/// </summary>
/// <remarks>
/// Before the recursive-descent parser was removed, this test compared the pratt parser's AST
/// against it as an oracle on every formula in all four datasets, and matched exactly (the
/// <c>minimumMatchingCount</c> below equalled each dataset's full "the oracle accepts it" count).
/// With no oracle left to compare against, this is now a smoke/regression test instead: for every
/// formula in the corpus,
/// <list type="bullet">
///   <item>a crash (an unrecognized failure reason - see <see cref="CategorizeFailure"/>) always
///   fails the test.</item>
///   <item>a plain rejection (a recognized "this construct isn't implemented" reason) is
///   expected and doesn't fail the test on its own - see the count below.</item>
/// </list>
/// The number of formulas the pratt parser currently handles is asserted as a floor, so this
/// catches a regression: raise the corresponding constant if a new, larger corpus (or new data
/// appended to an existing one) legitimately raises the count further; a drop means something
/// that used to parse no longer does.
/// </remarks>
public class PrattDataSetTests
{
    [Test]
    public async Task EnronDataSetCoverage()
    {
        // The remaining 17 gaps are confirmed not valid Excel syntax at all, not a pratt parser
        // gap: 15 formulas contain a truncated "#REF" error literal missing its "!" (e.g.
        // "SUM(#REF)"), and 2 use an unquoted sheet name containing spaces ("PJM Monthly Summary
        // 2000 08 V.1!..."), which A1's grammar requires quoting - both look like corruption from
        // however this corpus was originally extracted, not real formulas Excel ever accepted.
        await this.AssertCoverage("./data/enron/formulas.csv", minimumMatchingCount: 946_303);
    }

    [Test]
    public async Task EusesDataSetCoverage()
    {
        // The remaining gap is the same class of issue as enron's: the one failing formula
        // references a sheet ("Exercises 4, 5 and 6") whose name contains a comma and spaces but
        // isn't quoted with '...' as A1's grammar requires - not valid Excel syntax, not a pratt
        // parser gap.
        await this.AssertCoverage("./data/euses/formulas.csv", minimumMatchingCount: 89_294);
    }

    [Test]
    public async Task ContributionsDataSetCoverage()
    {
        await this.AssertCoverage("./data/contributions/formulas.csv", minimumMatchingCount: 1);
    }

    [Test]
    public async Task StructuredReferencesDataSetCoverage()
    {
        await this.AssertCoverage(
            "./data/structured-references/formulas.csv",
            minimumMatchingCount: 73
        );
    }

    private async Task AssertCoverage(string input, int minimumMatchingCount)
    {
        int consideredCount = 0;
        int matchingCount = 0;
        Dictionary<string, int> failureReasonCounts = [];
        List<string> unexpectedFailures = [];

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        foreach (string formula in DataSets.ReadCsv(input))
        {
            consideredCount++;

            try
            {
                parser.ParseFormula(formula, new Ctx());
                matchingCount++;
            }
            catch (Exception e)
            {
                string reason = CategorizeFailure(e);
                failureReasonCounts[reason] = failureReasonCounts.GetValueOrDefault(reason) + 1;
                if (
                    reason.StartsWith("Unexpected failure", StringComparison.Ordinal)
                    && unexpectedFailures.Count < 20
                )
                {
                    unexpectedFailures.Add($"'{formula}' => {e.GetType().Name}: {e.Message}");
                }
            }
        }

        TestContext.Current!.Output.WriteLine(
            $"{input}: pratt parser handles {matchingCount}/{consideredCount} "
                + $"({matchingCount * 100.0 / consideredCount:F1}%) of the formulas in this dataset."
        );
        foreach (
            KeyValuePair<string, int> reason in failureReasonCounts.OrderByDescending(kv =>
                kv.Value
            )
        )
        {
            TestContext.Current!.Output.WriteLine($"  {reason.Value, 8}  {reason.Key}");
        }

        int unexpectedFailureCount = failureReasonCounts
            .Where(kv => kv.Key.StartsWith("Unexpected failure", StringComparison.Ordinal))
            .Sum(kv => kv.Value);
        await Assert
            .That(unexpectedFailureCount)
            .IsEqualTo(0)
            .Because(
                "Pratt parser failed with an unrecognized error (a real bug, not a missing "
                    + $"feature) on:{Environment.NewLine}{string.Join(Environment.NewLine, unexpectedFailures)}"
            );

        await Assert
            .That(matchingCount)
            .IsGreaterThanOrEqualTo(minimumMatchingCount)
            .Because(
                $"Pratt parser coverage regressed: matched {matchingCount} formulas, expected at "
                    + $"least {minimumMatchingCount}. If this is an intentional improvement, raise "
                    + "the minimumMatchingCount baseline to lock in the progress."
            );
    }

    /// <summary>
    /// Bucket a parsing failure into a stable, human-readable reason. Buckets that describe a
    /// construct that simply isn't implemented yet are expected and don't fail the test; anything
    /// else (the "Unexpected failure" bucket) is a real bug and always fails the test - see
    /// <see cref="AssertCoverage"/>.
    /// </summary>
    private static string CategorizeFailure(Exception e)
    {
        const string missingParselet = "No parselet found for ";
        if (
            e is ParsingException
            && e.Message.Contains("wasn't fully consumed", StringComparison.Ordinal)
        )
        {
            return "Not implemented: leftover construct after the recognized prefix (e.g. a function call)";
        }

        if (
            e is InvalidOperationException
            && e.Message.StartsWith(missingParselet, StringComparison.Ordinal)
        )
        {
            string tokenType = e.Message[missingParselet.Length..].TrimEnd('.');
            return $"Not implemented: no parselet for token type {tokenType}";
        }

        if (
            e is ParsingException
            && e.Message.StartsWith("Expected token of type", StringComparison.Ordinal)
        )
        {
            return "Not implemented: unexpected token while parsing a partially-supported construct";
        }

        if (
            e is ParsingException
            && e.Message.StartsWith(
                "Unable to parse value starting from position",
                StringComparison.Ordinal
            )
        )
        {
            return "Not implemented: identifier doesn't match a currently recognized production";
        }

        if (e is ParsingException)
        {
            // Everything else thrown as a plain ParsingException comes from the lexer (e.g.
            // ParsingException.UnableToSelectToken for a unicode character class the pratt lexer
            // doesn't recognize as a valid identifier character) - a lexer gap, not a parser bug.
            return "Not implemented: lexer couldn't tokenize the formula";
        }

        return $"Unexpected failure: {e.GetType().Name}: {e.Message}";
    }
}
