using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests;

/// <summary>
/// R1C1-mode counterpart to <see cref="PrattDataSetTests"/>. There's no real-world R1C1 formula
/// corpus to draw from (every dataset under <c>data/</c> is A1), so instead each A1 formula the
/// oracle accepts is converted to R1C1 text via <see cref="FormulaConverter.ToR1C1"/> - the same
/// A1-to-R1C1 rewriter the rest of the codebase relies on - and *that* converted text is what both
/// the oracle (<see cref="FormulaParser{TScalarValue,TNode,TContext}.CellFormulaR1C1"/>) and the
/// pratt parser parse and compare, giving the same "must match exactly" differential guarantee the
/// A1 tests have, over real-world-shaped reference patterns rather than hand-written ones.
/// </summary>
public class PrattR1C1DataSetTests
{
    [Test]
    public async Task EnronDataSetCoverage()
    {
        await this.AssertCoverage("./data/enron/formulas.csv", minimumMatchingCount: 942_811);
    }

    private async Task AssertCoverage(string input, int minimumMatchingCount)
    {
        int consideredCount = 0;
        int matchingCount = 0;
        Dictionary<string, int> failureReasonCounts = [];
        List<string> unexpectedFailures = [];

        foreach (string formulaA1 in DataSets.ReadCsv(input))
        {
            string formula;
            try
            {
                // Not accepted by the oracle in A1 form, or the A1-to-R1C1 rewrite itself failed
                // (e.g. a relative reference that goes out of sheet bounds relative to the anchor)
                // - neither is this test's job to track; DataSetTests/PrattDataSetTests already do.
                formula = FormulaConverter.ToR1C1(formulaA1, row: 1, col: 1);
            }
            catch
            {
                continue;
            }

            AstNode oracleNode;
            try
            {
                oracleNode = FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaR1C1(
                    formula,
                    new Ctx(),
                    new F()
                );
            }
            catch
            {
                // The oracle itself rejects its own R1C1 rewrite of a formula it accepted in A1
                // form - a gap between the two oracle entry points, not this test's job to track.
                continue;
            }

            consideredCount++;

            AstNode prattNode;
            try
            {
                Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
                prattNode = parser.ParseFormula(formula, new Ctx(), isR1C1: true);
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
                    unexpectedFailures.Add(
                        $"'{formulaA1}' => R1C1 '{formula}' => {e.GetType().Name}: {e.Message}"
                    );
                }

                continue;
            }

            // A wrong result is worse than a thrown exception: fail immediately instead of
            // tallying, so a mismatch is never buried in the coverage statistics below.
            await Assert
                .That(prattNode)
                .IsEqualTo(oracleNode)
                .Because(
                    $"Pratt and oracle produced different ASTs for R1C1 '{formula}' (from A1 '{formulaA1}')"
                );
            matchingCount++;
        }

        TestContext.Current!.Output.WriteLine(
            consideredCount == 0
                ? $"{input}: oracle didn't accept any R1C1-converted formula from this dataset."
                : $"{input}: pratt parser matches the oracle on {matchingCount}/{consideredCount} "
                    + $"({matchingCount * 100.0 / consideredCount:F1}%) of the R1C1-converted formulas the oracle accepts."
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
                $"Pratt R1C1 coverage regressed: matched {matchingCount} formulas, expected at "
                    + $"least {minimumMatchingCount}. If this is an intentional improvement, raise "
                    + "the minimumMatchingCount baseline to lock in the progress."
            );
    }

    /// <summary>
    /// Same bucketing as <see cref="PrattDataSetTests.CategorizeFailure"/> - see there for the
    /// rationale of each bucket.
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
            return "Not implemented: lexer couldn't tokenize the formula";
        }

        return $"Unexpected failure: {e.GetType().Name}: {e.Message}";
    }
}
