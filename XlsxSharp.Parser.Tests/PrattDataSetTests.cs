using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests;

/// <summary>
/// The pratt parser (<see cref="XlsxSharp.Parser.Pratt"/>) is a work in progress replacement for
/// <see cref="FormulaParser{TScalarValue,TNode,TContext}"/>. These tests measure its progress
/// against the same real-world formula corpora used by <see cref="DataSetTests"/>, using the
/// existing recursive-descent parser as the oracle for what a correct result looks like.
/// </summary>
/// <remarks>
/// Unlike <see cref="DataSetTests"/>, these tests don't expect full coverage yet - that's the
/// whole point of the pratt parser rewrite still being in progress. Instead, for every formula
/// the oracle accepts:
/// <list type="bullet">
///   <item>If the pratt parser also accepts it, its AST must be identical to the oracle's. A
///   mismatch here is always a bug (a wrong result is worse than a thrown exception), so it fails
///   the test immediately.</item>
///   <item>If the pratt parser rejects it, the failure reason is bucketed. An unrecognized
///   failure reason (anything that isn't a plain "this construct isn't implemented yet" error,
///   e.g. a crash) fails the test, since that's a real bug rather than a missing feature.</item>
/// </list>
/// The number of formulas the pratt parser currently handles correctly is asserted as a floor:
/// this is the target picture (Zielbild) for the rewrite - the count should only ever go up as
/// features are implemented. Raise the corresponding constant whenever it does, so the test
/// documents current progress and catches regressions.
/// </remarks>
public class PrattDataSetTests
{
    [Test]
    public async Task EnronDataSetCoverage()
    {
        await this.AssertCoverage("./data/enron/formulas.csv", minimumMatchingCount: 752_730);
    }

    [Test]
    public async Task EusesDataSetCoverage()
    {
        await this.AssertCoverage("./data/euses/formulas.csv", minimumMatchingCount: 76_045);
    }

    [Test]
    public async Task ContributionsDataSetCoverage()
    {
        await this.AssertCoverage("./data/contributions/formulas.csv", minimumMatchingCount: 0);
    }

    [Test]
    public async Task StructuredReferencesDataSetCoverage()
    {
        await this.AssertCoverage(
            "./data/structured-references/formulas.csv",
            minimumMatchingCount: 1
        );
    }

    private async Task AssertCoverage(string input, int minimumMatchingCount)
    {
        int consideredCount = 0;
        int matchingCount = 0;
        Dictionary<string, int> failureReasonCounts = [];
        List<string> unexpectedFailures = [];

        foreach (string formula in DataSets.ReadCsv(input))
        {
            AstNode oracleNode;
            try
            {
                oracleNode = FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(
                    formula,
                    new Ctx(),
                    new F()
                );
            }
            catch
            {
                // Not accepted by our own reference implementation either - DataSetTests is
                // responsible for tracking that gap, not this one.
                continue;
            }

            consideredCount++;

            AstNode prattNode;
            try
            {
                Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
                prattNode = parser.ParseFormula(formula, new Ctx());
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

                continue;
            }

            // A wrong result is worse than a thrown exception: fail immediately instead of
            // tallying, so a mismatch is never buried in the coverage statistics below.
            await Assert
                .That(prattNode)
                .IsEqualTo(oracleNode)
                .Because($"Pratt and oracle produced different ASTs for '{formula}'");
            matchingCount++;
        }

        TestContext.Current!.Output.WriteLine(
            consideredCount == 0
                ? $"{input}: oracle didn't accept any formula from this dataset."
                : $"{input}: pratt parser matches the oracle on {matchingCount}/{consideredCount} "
                    + $"({matchingCount * 100.0 / consideredCount:F1}%) of the formulas the oracle accepts."
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
            e is InvalidOperationException
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
            // doesn't recognize as a valid identifier character, even though the Rolex-based
            // lexer used by the oracle does) - a lexer gap, not a parser bug.
            return "Not implemented: lexer couldn't tokenize the formula";
        }

        return $"Unexpected failure: {e.GetType().Name}: {e.Message}";
    }
}
