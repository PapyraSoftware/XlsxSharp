using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class StatisticalTests
{
    private const double Tolerance = 1e-6;
    private XLWorkbook workbook;

    [Test]
    public void Average()
    {
        double value;
        value = (double)this.workbook.Evaluate("AVERAGE(-27.5,93.93,64.51,-70.56)");
        ClassicAssert.AreEqual(15.095, value, Tolerance);

        IXLWorksheet ws = this.workbook.Worksheets.First();
        value = (double)ws.Evaluate("AVERAGE(G3:G45)");
        ClassicAssert.AreEqual(49.3255814, value, Tolerance);

        // Column D contains only strings - no average, because non-number types are skipped
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("AVERAGE(D3:D45)"));

        // Non-numbers in array are skipped instead of being converted
        ClassicAssert.AreEqual(-1, ws.Evaluate("AVERAGE({FALSE, TRUE, \"1\", \"0 0/2\", -1})"));

        // Blank value in references are skipped
        ws.Cell("Z1").Value = Blank.Value;
        ClassicAssert.AreEqual(1, ws.Evaluate("AVERAGE(Z1,1)"));

        AssertScalarToNumberConversion("AVERAGE", 0.5);
        AssertAnyErrorIsPropagated("AVERAGE");
    }

    [Test]
    public void AverageA()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Examples from specification
        ws.Cell("E1").Value = Blank.Value;
        ClassicAssert.AreEqual(10, ws.Evaluate("AVERAGEA(10, E1)"));
        ws.Cell("E2").Value = true;
        ClassicAssert.AreEqual(5.5, ws.Evaluate("AVERAGEA(10, E2)"));
        ws.Cell("E3").Value = false;
        ClassicAssert.AreEqual(5, ws.Evaluate("AVERAGEA(10, E3)"));

        // Make sure multiple values not in an array work as intended
        ClassicAssert.AreEqual(
            15.095,
            (double)this.workbook.Evaluate("AVERAGEA(-27.5,93.93,64.51,-70.56)"),
            Tolerance
        );

        // Array logical arguments are ignored
        ClassicAssert.AreEqual(2, this.workbook.Evaluate("AVERAGEA({2,TRUE,TRUE,FALSE,FALSE})"));

        // Array text arguments are counted as zero (4+2+0+0)/4
        ClassicAssert.AreEqual(1.5, this.workbook.Evaluate("AVERAGEA({4, 2, \"hello\", \"10\" })"));

        // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
        ws.Cell("Z1").Value = Blank.Value; // Not counted
        ws.Cell("Z2").Value = true; // 1
        ws.Cell("Z3").Value = "100"; // 0
        ws.Cell("Z4").Value = "hello"; // 0
        ws.Cell("Z5").Value = 0; // 0
        ws.Cell("Z6").Value = 4; // 4
        ClassicAssert.AreEqual(1, (double)ws.Evaluate("AVERAGEA(Z1:Z6)"));

        AssertScalarToNumberConversion("AVERAGEA", 0.5);
        AssertAnyErrorIsPropagated("AVERAGEA");
    }

    [Test]
    [Arguments(6, 10, 0.5, 0.205078125)]
    [Arguments(4, 20, 0.2, 0.2181994)] // p different than 0.5
    [Arguments(0, 5, 0.2, 0.32768)] // 0 out of 5 successes
    [Arguments(0, 0, 0.2, 1)] // 0 out of 0 successes
    [Arguments(1, 1, 0, 0)]
    [Arguments(1, 1, 1, 1)]
    [Arguments(2, 4, 0.5, 0.375)]
    [Arguments(2.9, 4.9, 0.5, 0.375)] // Attempts are floored
    public void BinomDistCalculatesNonCumulativeBinomialDistribution(
        double k,
        double n,
        double p,
        double expected
    )
    {
        string kString = k.ToInvariantString();
        string nString = n.ToInvariantString();
        string pString = p.ToInvariantString();
        double result = (double)
            XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, FALSE)");
        ClassicAssert.AreEqual(expected, result, Tolerance);
    }

    [Test]
    [Arguments(6, 10, 0.5, 0.828125)]
    [Arguments(2, 7, 0.3, 0.6470695)]
    [Arguments(0, 7, 0.3, 0.0823543)]
    [Arguments(0, 0, 0.3, 1)]
    [Arguments(0, 0, 1, 1)]
    [Arguments(2, 4, 0.5, 0.6875)]
    [Arguments(2.9, 4.9, 0.5, 0.6875)] // Values are floored
    public void BinomDistCalculatesCumulativeBinomialDistribution(
        double k,
        double n,
        double p,
        double expected
    )
    {
        string kString = k.ToInvariantString();
        string nString = n.ToInvariantString();
        string pString = p.ToInvariantString();
        double result = (double)
            XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, TRUE)");
        ClassicAssert.AreEqual(expected, result, Tolerance);
    }

    [Test]
    [Arguments(5, 4, 0.5)] // Five successes out of 4 attempts
    [Arguments(-1, 4, 0.5)] // Negative successes
    [Arguments(0, -1, 0.5)] // Negative attempts
    [Arguments(2, 4, -0.1)] // p < 0
    [Arguments(2, 4, 1.1)] // p > 1
    [Arguments(1E+300, 2E+300, 0.5)] // Too large values
    public void BinomDistReturnsNumErrorOnInvalidCalculations(double k, double n, double p)
    {
        string kString = k.ToInvariantString();
        string nString = n.ToInvariantString();
        string pString = p.ToInvariantString();
        XLCellValue result = XLWorkbook.EvaluateExpr(
            $"BINOMDIST({kString}, {nString}, {pString}, FALSE)"
        );
        ClassicAssert.AreEqual(XLError.NumberInvalid, result);
    }

    [Test]
    public void Count()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate("COUNT(D3:D45)");
        ClassicAssert.AreEqual(0, value);

        value = ws.Evaluate("COUNT(G3:G45)");
        ClassicAssert.AreEqual(43, value);

        value = ws.Evaluate("COUNT(G:G)");
        ClassicAssert.AreEqual(43, value);

        value = this.workbook.Evaluate("COUNT(Data!G:G)");
        ClassicAssert.AreEqual(43, value);

        // Scalar blank, logical and text is counted as numbers
        ClassicAssert.AreEqual(4, ws.Evaluate("COUNT(IF(TRUE,,),TRUE, FALSE, \"1\")"));

        // Non-number values in arrays are not counted as numbers.
        ClassicAssert.AreEqual(0, ws.Evaluate("COUNT({TRUE,FALSE,\"1\"})"));

        // Text is not counted as number.
        ClassicAssert.AreEqual(0, ws.Evaluate("COUNT(\"Hello\")"));

        // Blank cells are not counted as numbers
        ws.Cell("Z1").Value = Blank.Value;
        ClassicAssert.AreEqual(0, ws.Evaluate("COUNT(Z1)"));

        // Scalar errors are not propagated
        ClassicAssert.AreEqual(1, ws.Evaluate("COUNT(1, #NULL!)"));

        // Array errors are not propagated
        ClassicAssert.AreEqual(1, ws.Evaluate("COUNT({1, #NULL!})"));

        // Reference errors are not propagated
        ws.Cell("Z1").Value = XLError.NullValue;
        ClassicAssert.AreEqual(0, ws.Evaluate("COUNT(Z1)"));
    }

    [Test]
    public void CountA()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();
        XLCellValue value = ws.Evaluate("COUNTA(D3:D45)");
        ClassicAssert.AreEqual(43, value);

        value = ws.Evaluate("COUNTA(G3:G45)");
        ClassicAssert.AreEqual(43, value);

        value = ws.Evaluate("COUNTA(G:G)");
        ClassicAssert.AreEqual(44, value);

        value = this.workbook.Evaluate("COUNTA(Data!G:G)");
        ClassicAssert.AreEqual(44, value);
    }

    [Test]
    public void CountACountsNonBlankValues()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = 39790;
        ws.Cell("A3").Value = 0;
        ws.Cell("A4").Value = 22.24;
        ws.Cell("A5").Value = "Text";
        ws.Cell("A6").Value = false;
        ws.Cell("A7").Value = true;
        ws.Cell("A8").Value = XLError.DivisionByZero;
        ws.Cell("A9").FormulaA1 = "COUNTA(A1:B8)";
        ClassicAssert.AreEqual(7, ws.Cell("A9").Value);
    }

    [Test]
    public void CountAOnExamplesFromSpec()
    {
        ClassicAssert.AreEqual(5, XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)"));
        ClassicAssert.AreEqual(5, XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)"));
        ClassicAssert.AreEqual(7, XLWorkbook.EvaluateExpr("COUNTA({1,2,3,4,5},6,\"7\")"));

        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("E2").Value = true;
        ClassicAssert.AreEqual(1, ws.Evaluate("COUNTA(10, E1)"));
        ClassicAssert.AreEqual(2, ws.Evaluate("COUNTA(10, E2)"));
    }

    [Test]
    public void CountAAcceptsUnionReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A2").Value = 7;
        ws.Cell("B5").Value = false;
        ClassicAssert.AreEqual(2, ws.Evaluate("COUNTA((A1:A4,B4:B7))"));
    }

    [Test]
    public void CountADoesntCountSingleBlankCellReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(0, ws.Evaluate("COUNTA(A1)"));
    }

    [Test]
    public void CountACountsBlankArgument() =>
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("COUNTA(IF(TRUE,,))"));

    [Test]
    public void CountACountsErrorArguments() =>
        ClassicAssert.AreEqual(
            7,
            XLWorkbook.EvaluateExpr("COUNTA(#NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A)")
        );

    [Test]
    public void CountACountsEmptyString()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = string.Empty;
        ClassicAssert.AreEqual(2, ws.Evaluate("COUNTA(A1, \"\")"));
    }

    [Test]
    public void CountBlank()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = 0;
        ws.Cell("A3").Value = 1;
        ws.Cell("A4").Value = false;
        ws.Cell("A5").Value = true;
        ws.Cell("A6").Value = "";
        ws.Cell("A7").Value = "Text";
        ws.Cell("A8").Value = XLError.DivisionByZero;

        // Blank and empty text value is counted as blank
        ClassicAssert.AreEqual(1, ws.Evaluate("COUNTBLANK(A1)"));
        ClassicAssert.AreEqual(string.Empty, ws.Cell("A6").Value);
        ClassicAssert.AreEqual(1, ws.Evaluate("COUNTBLANK(A6)"));

        // Anything else isn't counted as blank
        ClassicAssert.AreEqual(2, ws.Evaluate("COUNTBLANK(A1:A8)"));

        ClassicAssert.AreEqual(17179869178d, ws.Evaluate("COUNTBLANK(A:XFD)"));

        // Check that all others argument types. The Excel grammar doesn't allow that,
        // so use IF workaround for that.
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("COUNTBLANK(IF(TRUE,))")); // Blank
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            ws.Evaluate("COUNTBLANK(IF(TRUE,FALSE))")
        ); // Logical
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("COUNTBLANK(IF(TRUE,1))")); // Number
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("COUNTBLANK(IF(TRUE,\"\"))")); // Text
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("COUNTBLANK(IF(TRUE,#DIV/0!))")); // Error
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("COUNTBLANK(IF(TRUE,{1}))")); // Array
    }

    [Test]
    public void CountIf()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate(@"=COUNTIF(D3:D45,""Central"")");
        ClassicAssert.AreEqual(24, value);

        value = ws.Evaluate(@"=COUNTIF(D:D,""Central"")");
        ClassicAssert.AreEqual(24, value);

        value = this.workbook.Evaluate(@"=COUNTIF(Data!D:D,""Central"")");
        ClassicAssert.AreEqual(24, value);
    }

    [Test]
    [Arguments(@"=COUNTIF(Data!E:E, ""J*"")", 13)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*i*"")", 21)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*in*"")", 9)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*i*l"")", 9)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*i?e*"")", 9)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*o??s*"")", 10)]
    [Arguments(@"=COUNTIF(Data!X1:X1000, """")", 1000)]
    [Arguments(@"=COUNTIF(Data!E1:E44, """")", 1)]
    public void CountIfConditionWithWildcards(string formula, int expectedResult)
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        XLCellValue value = ws.Evaluate(formula);
        ClassicAssert.AreEqual(expectedResult, value);
    }

    [Test]
    [Arguments(@"=COUNTIF(A1:A10, 1)", 1)]
    [Arguments(@"=COUNTIF(A1:A10, 2.0)", 1)]
    [Arguments(@"=COUNTIF(A1:A10, ""3"")", 2)]
    [Arguments(@"=COUNTIF(A1:A10, 3)", 2)]
    [Arguments(@"=COUNTIF(A1:A10, 43831)", 1)]
    [Arguments(@"=COUNTIF(A1:A10, DATE(2020, 1, 1))", 1)]
    [Arguments(@"=COUNTIF(A1:A10, TRUE)", 1)]
    public void CountIfMixedData(string formula, int expected)
    {
        // We follow to Excel's convention.
        // Excel treats 1 and TRUE as unequal, but 3 and "3" as equal
        // LibreOffice Calc handles some SUMIF and COUNTIF differently, e.g. it treats 1 and TRUE as equal, but 3 and "3" differently
        IXLWorksheet ws = this.workbook.Worksheet("MixedData");
        ClassicAssert.AreEqual(expected, ws.Evaluate(formula));
    }

    [Test]
    [Arguments("x", @"=COUNTIF(A1:A1, ""?"")", 1)]
    [Arguments("x", @"=COUNTIF(A1:A1, ""~?"")", 0)]
    [Arguments("?", @"=COUNTIF(A1:A1, ""~?"")", 1)]
    [Arguments("~?", @"=COUNTIF(A1:A1, ""~?"")", 0)]
    [Arguments("~?", @"=COUNTIF(A1:A1, ""~~~?"")", 1)]
    [Arguments("?", @"=COUNTIF(A1:A1, ""~~?"")", 0)]
    [Arguments("~?", @"=COUNTIF(A1:A1, ""~~?"")", 1)]
    [Arguments("~x", @"=COUNTIF(A1:A1, ""~~?"")", 1)]
    [Arguments("*", @"=COUNTIF(A1:A1, ""~*"")", 1)]
    [Arguments("~*", @"=COUNTIF(A1:A1, ""~*"")", 0)]
    [Arguments("~*", @"=COUNTIF(A1:A1, ""~~~*"")", 1)]
    [Arguments("*", @"=COUNTIF(A1:A1, ""~~*"")", 0)]
    [Arguments("~*", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
    [Arguments("~x", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
    [Arguments("~xyz", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
    public void CountIfMoreWildcards(string cellContent, string formula, int expectedResult)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.Cell(1, 1).Value = cellContent;

            ClassicAssert.AreEqual(expectedResult, (double)ws.Evaluate(formula));
        }
    }

    [Test]
    [Arguments("=COUNTIFS(B1:D1, \"=Yes\")", 1)]
    [Arguments("=COUNTIFS(B1:B4, \"=Yes\", C1:C4, \"=Yes\")", 2)]
    [Arguments("=COUNTIFS(B4:D4, \"=Yes\", B2:D2, \"=Yes\")", 1)]
    public void CountIfsReferenceExample1FromExcelDocumentations(
        string formula,
        int expectedOutcome
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.Cell(1, 1).Value = "Davidoski";
            ws.Cell(1, 2).Value = "Yes";
            ws.Cell(1, 3).Value = "No";
            ws.Cell(1, 4).Value = "No";

            ws.Cell(2, 1).Value = "Burke";
            ws.Cell(2, 2).Value = "Yes";
            ws.Cell(2, 3).Value = "Yes";
            ws.Cell(2, 4).Value = "No";

            ws.Cell(3, 1).Value = "Sundaram";
            ws.Cell(3, 2).Value = "Yes";
            ws.Cell(3, 3).Value = "Yes";
            ws.Cell(3, 4).Value = "Yes";

            ws.Cell(4, 1).Value = "Levitan";
            ws.Cell(4, 2).Value = "No";
            ws.Cell(4, 3).Value = "Yes";
            ws.Cell(4, 4).Value = "Yes";

            ClassicAssert.AreEqual(expectedOutcome, ws.Evaluate(formula));
        }
    }

    [Test]
    public void CountIfsSingleCondition()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate(@"=COUNTIFS(D3:D45,""Central"")");
        ClassicAssert.AreEqual(24, value);

        value = ws.Evaluate(@"=COUNTIFS(D:D,""Central"")");
        ClassicAssert.AreEqual(24, value);

        value = this.workbook.Evaluate(@"=COUNTIFS(Data!D:D,""Central"")");
        ClassicAssert.AreEqual(24, value);
    }

    [Test]
    [Arguments(@"=COUNTIFS(Data!E:E, ""J*"")", 13)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*i*"")", 21)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*in*"")", 9)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*i*l"")", 9)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*i?e*"")", 9)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*o??s*"")", 10)]
    [Arguments(@"=COUNTIFS(Data!X1:X1000, """")", 1000)]
    [Arguments(@"=COUNTIFS(Data!E1:E44, """")", 1)]
    public void CountIfsSingleConditionWithWildcards(string formula, int expectedResult)
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        XLCellValue value = ws.Evaluate(formula);
        ClassicAssert.AreEqual(expectedResult, value);
    }

    [Test]
    [Arguments("COUNTIFS(H1:I3, 1, D1:F2, 2)")]
    [Arguments("COUNTIFS(A:B, \"A*\", C:C, \">2\")")]
    public void CountIfsReturnsErrorWhenAreasDimensionsAreDifferent(string formula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula));
    }

    // TUnit creates a new instance of this class per test (unlike NUnit's one shared fixture
    // instance), so there's no single "last" workbook to dispose once at the end. Each test's own
    // workbook is disposed right after that test instead.
    [After(Test)]
    public void Dispose() => this.workbook.Dispose();

    [Test]
    [Arguments("H3:H45", 7.51126069234216)]
    [Arguments("H:H", 7.51126069234216)]
    [Arguments("Data!H:H", 7.51126069234216)]
    [Arguments("H3:H10", 5.26214814727941)]
    [Arguments("H3:H20", 7.01281435054797)]
    [Arguments("H3:H30", 7.00137389296182)]
    [Arguments("H3:H3", 1.99)]
    [Arguments("H10:H20", 8.37855107505682)]
    [Arguments("H15:H20", 15.8927310267677)]
    [Arguments("H20:H30", 7.14321227391814)]
    public void GeomeanCalculation(string sourceValue, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)this.workbook.Worksheets.First().Evaluate($"GEOMEAN({sourceValue})"),
            1e-12
        );
    }

    [Test]
    [Arguments("D3:D45", XLError.NumberInvalid)]
    [Arguments("-1, 0, 3", XLError.NumberInvalid)]
    [Arguments("0", XLError.NumberInvalid)]
    public void GeomeanIncorrectCases(string sourceValue, XLError expected)
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        ClassicAssert.AreEqual(expected, (XLError)ws.Evaluate($"GEOMEAN({sourceValue})"));
    }

    [Test]
    public void Geomean()
    {
        // Example from the specification
        ClassicAssert.AreEqual(
            5.4444547024966,
            (double)XLWorkbook.EvaluateExpr("GEOMEAN(10.5,5.3,2.9)"),
            1e-8
        );
        ClassicAssert.AreEqual(
            6.6337805880630,
            (double)XLWorkbook.EvaluateExpr("GEOMEAN(10.5,{5.3,2.9},\"12\")"),
            1e-8
        );

        // GEOMEAN isn't limited by double scale, i.e. it doesn't use naive algorithm for large number.
        ClassicAssert.AreEqual(
            1.0000000000000231E+307d,
            (double)XLWorkbook.EvaluateExpr("GEOMEAN(1E+307, 1E+307)"),
            1e-8
        );

        // Scalar blank is counted as a 0
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("GEOMEAN(IF(TRUE,), 1)")
        );

        // Scalar logical and text is converted to numbers
        ClassicAssert.AreEqual(
            2.236067977,
            (double)XLWorkbook.EvaluateExpr("GEOMEAN(TRUE, \"5\")"),
            1e-8
        );

        // Non-number values in arrays are ignored.
        ClassicAssert.AreEqual(
            5.916079783,
            (double)XLWorkbook.EvaluateExpr("GEOMEAN({TRUE, FALSE, \"1\", 7}, 5)"),
            1e-8
        );

        // Scalar non-number text causes an error due to conversion.
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("GEOMEAN(\"Hello\", 5)")
        );

        // Reference non-number arguments are ignored
        IXLWorksheet ws = this.workbook.Worksheets.First();
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = "1";
        ws.Cell("Z3").Value = "hello";
        ws.Cell("Z4").Value = false;
        ws.Cell("Z5").Value = true;
        ws.Cell("Z6").Value = 5;
        ClassicAssert.AreEqual(5, (double)ws.Evaluate("GEOMEAN(Z1:Z6)"), 1e-8);

        AssertAnyErrorIsPropagated("GEOMEAN");
    }

    [Before(Test)]
    public void Init() => this.workbook = SetupWorkbook();

    [Test]
    [Arguments(@"H3:H45", 94145.5271162791)]
    [Arguments(@"H:H", 94145.5271162791)]
    [Arguments(@"Data!H:H", 94145.5271162791)]
    [Arguments(@"H3:H10", 411.5)]
    [Arguments(@"H3:H20", 13604.2067611111)]
    [Arguments(@"H3:H30", 14231.0694)]
    [Arguments(@"H3:H3", 0)]
    [Arguments(@"H10:H20", 12713.7600909091)]
    [Arguments(@"H15:H20", 10827.2200833333)]
    [Arguments(@"H20:H30", 477.132272727273)]
    public void DevSq(string sourceValue, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)this.workbook.Worksheets.First().Evaluate($"DEVSQ({sourceValue})"),
            1e-10
        );
    }

    [Test]
    [Arguments("D3:D45", XLError.NumberInvalid)]
    public void DevsqIncorrectCases(string sourceValue, XLError expected)
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        ClassicAssert.AreEqual(expected, (XLError)ws.Evaluate($"DEVSQ({sourceValue})"));
    }

    [Test]
    public void DevsqIsCalculatedFromNumbers()
    {
        ClassicAssert.AreEqual(
            6.90666666666666,
            (double)XLWorkbook.EvaluateExpr("DEVSQ(5.6, 8.2, 9.2)"),
            1e-10
        );
        ClassicAssert.AreEqual(
            6.90666666666666,
            (double)XLWorkbook.EvaluateExpr("DEVSQ({ 5.6, 8.2, 9.2})"),
            1e-10
        );

        // Array logical arguments are ignored
        ClassicAssert.AreEqual(
            0,
            (double)this.workbook.Evaluate("DEVSQ({2,TRUE,TRUE,FALSE,FALSE})"),
            1e-10
        );
        ClassicAssert.AreEqual(
            2.8,
            (double)this.workbook.Evaluate("DEVSQ({2, 1, 1, 0, 0})"),
            1e-10
        );

        // Array text arguments are ignored
        ClassicAssert.AreEqual(
            2,
            (double)this.workbook.Evaluate("DEVSQ({4, 2, \"hello\", \"10\" })"),
            1e-10
        );

        // Non-numerical reference values are ignored.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value; // Ignored
        ws.Cell("A2").Value = true; // Ignored
        ws.Cell("A3").Value = "100"; // Ignored
        ws.Cell("A4").Value = "hello"; // Ignored
        ws.Cell("A5").Value = 2; // Included
        ws.Cell("A6").Value = 4; // Included
        ClassicAssert.AreEqual(2, (double)ws.Evaluate("DEVSQ(A1:A6)"), 1e-10);

        AssertScalarToNumberConversion("DEVSQ", 0.5);
        AssertAnyErrorIsPropagated("DEVSQ");
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.2, 0.202732554054082)]
    [Arguments(0.25, 0.255412811882995)]
    [Arguments(0.3296001056, 0.342379555936801)]
    [Arguments(-0.36, -0.37688590118819)]
    [Arguments(-0.000003, -0.00000299999999998981)]
    [Arguments(-0.063453535345348, -0.0635389037459617)]
    [Arguments(0.559015883901589171354964, 0.631400600322212)]
    [Arguments(0.2691496, 0.275946780611959)]
    [Arguments(-0.10674142, -0.107149608461448)]
    public void Fisher(double sourceValue, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"FISHER({sourceValue})"),
            1e-12
        );
    }

    [Test]
    [Arguments("\"asdf\"", XLError.IncompatibleValue)]
    [Arguments("5", XLError.NumberInvalid)]
    [Arguments("-1", XLError.NumberInvalid)]
    [Arguments("1", XLError.NumberInvalid)]
    public void FisherIncorrectCases(string sourceValue, XLError expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (XLError)XLWorkbook.EvaluateExpr($"FISHER({sourceValue})")
        );
    }

    [Test]
    public void Max()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate(@"=MAX(D3:D45)");
        ClassicAssert.AreEqual(0, value);

        value = ws.Evaluate(@"=MAX(G3:G45)");
        ClassicAssert.AreEqual(96, value);

        value = ws.Evaluate(@"=MAX(G:G)");
        ClassicAssert.AreEqual(96, value);

        value = this.workbook.Evaluate(@"=MAX(Data!G:G)");
        ClassicAssert.AreEqual(96, value);

        // Although in most cases blank cells are considered 0, MAX just ignores them.
        value = this.workbook.Evaluate(@"MAX(-10, Data!X:Z)");
        ClassicAssert.AreEqual(-10, value);

        // Arrays - numbers are used
        value = this.workbook.Evaluate(@"MAX(-10, { -6, -5, 7 })");
        ClassicAssert.AreEqual(7, value);

        // Arrays - non-number and non-error values are skipped.
        value = this.workbook.Evaluate(@"MAX(-10, { TRUE, FALSE, ""100"" })");
        ClassicAssert.AreEqual(-10, value);

        // Reference argument ignores everything but number.
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = -4;
        ClassicAssert.AreEqual(-4, ws.Evaluate("MAX(Z1:Z5)"));

        AssertScalarToNumberConversion("MAX", 1);
        AssertAnyErrorIsPropagated("MAX");
    }

    [Test]
    public void MaxA()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Examples from specification
        ClassicAssert.AreEqual(12.6, ws.Evaluate("MAXA(10.4,-3.5,12.6)"));
        ClassicAssert.AreEqual(12.6, ws.Evaluate("MAXA(10.4,{-3.5,12.6})"));
        ClassicAssert.AreEqual(0, ws.Evaluate("MAXA({\"ABC\",TRUE})"));
        ws.Cell("B3").Value = Blank.Value;
        ClassicAssert.AreEqual(-10, ws.Evaluate("MAX(-10,-12,-15,B3)"));
        ws.Cell("B3").Value = 0;
        ClassicAssert.AreEqual(0, ws.Evaluate("MAXA(-10,-12,-15,B3)"));

        // Array logical arguments are ignored
        ClassicAssert.AreEqual(-2, this.workbook.Evaluate("MAXA({-2, TRUE, TRUE, FALSE, FALSE})"));

        // Array text arguments are ignored
        ClassicAssert.AreEqual(-2, this.workbook.Evaluate("MAXA({-4, -2, \"hello\", \"10\" })"));

        // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = -4;
        ClassicAssert.AreEqual(1, ws.Evaluate("MAXA(A1:A5)"));
        ClassicAssert.AreEqual(0, ws.Evaluate("MAXA(A3:A5)"));

        AssertScalarToNumberConversion("MAXA", 1);
        AssertAnyErrorIsPropagated("MAXA");
    }

    [Test]
    public void MedianWithAreaWithoutNumericValuesReturnsError()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        // Column D contains names of regions
        ClassicAssert.AreEqual(XLError.NumberInvalid, ws.Evaluate("MEDIAN(D3:D45)"));
    }

    [Test]
    public void MedianEvenCountOfCellRangeReturnsAverageOfTwoElementsInMiddleOfSortedList()
    {
        //Arrange
        IXLWorksheet ws = this.workbook.Worksheets.First();

        //Act
        double value = (double)ws.Evaluate("MEDIAN(I3:I10)");

        //Assert
        ClassicAssert.AreEqual(244.225, value, Tolerance);
    }

    [Test]
    public void MedianEvenCountOfManualNumbersReturnsAverageOfTwoElementsInMiddleOfSortedList()
    {
        //Act
        double value = (double)this.workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56)");

        //Assert
        ClassicAssert.AreEqual(18.505, value, Tolerance);
    }

    [Test]
    public void MedianOddCountOfCellRangeReturnsElementInMiddleOfSortedList()
    {
        //Arrange
        IXLWorksheet ws = this.workbook.Worksheets.First();

        //Act
        double value = (double)ws.Evaluate("MEDIAN(I3:I11)");

        //Assert
        ClassicAssert.AreEqual(189.05, value, Tolerance);
    }

    [Test]
    public void MedianOddCountOfManualNumbersReturnsElementInMiddleOfSortedList()
    {
        //Act
        double value = (double)this.workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56,101.65)");

        //Assert
        ClassicAssert.AreEqual(64.51, value, Tolerance);
    }

    [Test]
    public void MedianUsesOnlyNumbers()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Examples from specification
        ClassicAssert.AreEqual(15, ws.Evaluate("MEDIAN(10, 20)"));
        ClassicAssert.AreEqual(-1.05, ws.Evaluate("MEDIAN(-3.5, 1.4, 6.9, -4.5)"));
        ClassicAssert.AreEqual(-1.05, ws.Evaluate("MEDIAN({ -3.5,1.4,6.9},-4.5)"));

        // Reference with no value will return error
        ws.Cell("A1").Value = Blank.Value;
        ClassicAssert.AreEqual(XLError.NumberInvalid, ws.Evaluate("MEDIAN(A1)"));

        // Array non-number values are ignored
        ClassicAssert.AreEqual(7, ws.Evaluate("MEDIAN({7, TRUE,FALSE,\"1\"})"));

        // Only numbers are used from reference, rest is ignored
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = 0;
        ws.Cell("A6").Value = 4;
        ws.Cell("A7").Value = 5;
        ClassicAssert.AreEqual(4, ws.Evaluate("MEDIAN(A1:A7)"));

        AssertScalarToNumberConversion("MEDIAN", 0.5);
        AssertAnyErrorIsPropagated("MEDIAN");
    }

    [Test]
    public void Min()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();
        ClassicAssert.AreEqual(0, ws.Evaluate("MIN(D3:D45)"));
        ClassicAssert.AreEqual(2, ws.Evaluate("MIN(G3:G45)"));
        ClassicAssert.AreEqual(2, ws.Evaluate("MIN(G:G)"));
        ClassicAssert.AreEqual(2, this.workbook.Evaluate("MIN(Data!G:G)"));

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            5,
            this.workbook.Evaluate("MIN({5, TRUE, FALSE, \"1\", \"hello\"})")
        );

        // Reference non-number arguments are ignored
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = "1";
        ws.Cell("Z3").Value = "hello";
        ws.Cell("Z4").Value = false;
        ws.Cell("Z5").Value = true;
        ws.Cell("Z6").Value = 5;
        ClassicAssert.AreEqual(5, ws.Evaluate("MIN(Z1:Z6)"));

        // If there is no value, return 0
        ClassicAssert.AreEqual(0, ws.Evaluate("MIN({\"hello\"})"));

        AssertScalarToNumberConversion("MIN", 0);
        AssertAnyErrorIsPropagated("MIN");
    }

    [Test]
    public void MinA()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Examples from specification
        ClassicAssert.AreEqual(-3.5, ws.Evaluate("MINA(10.4, -3.5, 12.6)"));
        ClassicAssert.AreEqual(-3.5, ws.Evaluate("MINA(10.4, {-3.5, 12.6})"));
        ClassicAssert.AreEqual(0, ws.Evaluate("MINA({\"ABC\", TRUE})"));
        ws.Cell("B3").Value = Blank.Value;
        ClassicAssert.AreEqual(10, ws.Evaluate("MINA(10, 12, 15, B3)"));
        ws.Cell("B3").Value = "Text";
        ClassicAssert.AreEqual(0, ws.Evaluate("MINA(10, 12, 15, B3)"));

        // Blanks in references are ignored and when MINA doesn't have any values, it returns 0
        ws.Cell("A1").Value = Blank.Value;
        ClassicAssert.AreEqual(0, ws.Evaluate("MINA(A1)"));

        // Array logical arguments are ignored
        ClassicAssert.AreEqual(2, wb.Evaluate("MINA({2, TRUE, TRUE, FALSE, FALSE})"));

        // Array text arguments are ignored
        ClassicAssert.AreEqual(2, wb.Evaluate("MINA({4, 2, \"hello\", \"1\"})"));

        // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
        ws.Cell("A1").Value = Blank.Value; // Ignores
        ws.Cell("A2").Value = true; // Includes
        ws.Cell("A3").Value = "100"; // Considers 0
        ws.Cell("A4").Value = "hello"; // Considers 0
        ws.Cell("A5").Value = -4; // Included
        ClassicAssert.AreEqual(1, ws.Evaluate("MINA(A1:A2)"));
        ClassicAssert.AreEqual(0, ws.Evaluate("MINA(A1:A3)"));
        ClassicAssert.AreEqual(-4, ws.Evaluate("MINA(A1:A5)"));

        AssertScalarToNumberConversion("MINA", 0);
        AssertAnyErrorIsPropagated("MINA");
    }

    [Test]
    public void StDev()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        // Only non-convertible text in D column, thus less than 2 samples will return error
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("STDEV(D3:D45)"));

        // Calculate StDev from numeric values (reference contains only numbers)
        double value = (double)ws.Evaluate("STDEV(H3:H45)");
        ClassicAssert.AreEqual(47.34511769, value, Tolerance);

        // Ignores text values in the H column and only uses numeric ones, same as reference with only number
        value = (double)ws.Evaluate("STDEV(H:H)");
        ClassicAssert.AreEqual(47.34511769, value, Tolerance);

        value = (double)this.workbook.Evaluate("STDEV(Data!H:H)");
        ClassicAssert.AreEqual(47.34511769, value, Tolerance);

        // Need at least two values, otherwise returns error
        ClassicAssert.AreEqual(XLError.DivisionByZero, this.workbook.Evaluate("STDEV(1)"));
        ClassicAssert.AreEqual(0, (double)this.workbook.Evaluate("STDEV(0, 0)"), Tolerance);

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            0.707106781,
            (double)this.workbook.Evaluate("STDEV({0, 1, \"Hello\", FALSE, TRUE})"),
            Tolerance
        );

        // Reference argument only uses number, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        ClassicAssert.AreEqual(0.707106781, (double)ws.Evaluate("STDEV(Z1:Z6)"), Tolerance);

        AssertScalarToNumberConversion("STDEV", 0.707106781);
        AssertAnyErrorIsPropagated("STDEV");
    }

    [Test]
    public void StDevA()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Example from specification
        ClassicAssert.AreEqual(
            23.72902583,
            (double)ws.Evaluate("STDEVA(123, 134, 143, 173, 112, 109)"),
            Tolerance
        );

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            0.707106781,
            (double)ws.Evaluate("STDEVA({0, 1, \"9\", \"Hello\", FALSE, TRUE})"),
            Tolerance
        );

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        ClassicAssert.AreEqual(3.060501048, (double)ws.Evaluate("STDEVA(A1:A7)"), Tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("STDEVA({\"hello\"})"));

        AssertScalarToNumberConversion("STDEVA", 0.707106781);
        AssertAnyErrorIsPropagated("STDEVA");
    }

    [Test]
    public void StDevP()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        // Example from specification
        ClassicAssert.AreEqual(
            21.66153785,
            (double)ws.Evaluate("STDEVP(123, 134, 143, 173, 112, 109)"),
            Tolerance
        );

        // Column D contains only region names (non-convertible text), thus reference contains less than 1 sample that is required
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("STDEVP(D3:D45)"));

        // Calculate StDevP from numeric values (reference contains only numbers)
        ClassicAssert.AreEqual(46.79135458, (double)ws.Evaluate("STDEVP(H3:H45)"), Tolerance);

        // StDevP ignores text values/blanks in the H column and only uses numeric ones, the result is same as the reference above that contains only numbers
        ClassicAssert.AreEqual(46.79135458, (double)ws.Evaluate("STDEVP(H:H)"), Tolerance);

        ClassicAssert.AreEqual(
            46.79135458,
            (double)this.workbook.Evaluate("STDEVP(Data!H:H)"),
            Tolerance
        );

        // If sample size is 0, return error
        ClassicAssert.AreEqual(XLError.DivisionByZero, this.workbook.Evaluate("STDEVP({TRUE})"));
        ClassicAssert.AreEqual(0, this.workbook.Evaluate("STDEVP(100)"));

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            0.5,
            this.workbook.Evaluate("STDEVP({0, 1, \"Hello\", FALSE, TRUE})")
        );

        // Reference argument only uses numbers, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        ClassicAssert.AreEqual(0.5, ws.Evaluate("STDEVP(Z1:Z6)"));

        AssertScalarToNumberConversion("STDEVP", 0.5);
        AssertAnyErrorIsPropagated("STDEVP");
    }

    [Test]
    public void StDevPa()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Example from specification
        ClassicAssert.AreEqual(
            21.66153785,
            (double)ws.Evaluate("STDEVPA(123, 134, 143, 173, 112, 109)"),
            Tolerance
        );

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            0.5,
            (double)ws.Evaluate("STDEVPA({0, 1, \"9\", \"Hello\", FALSE, TRUE})"),
            Tolerance
        );

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        ClassicAssert.AreEqual(2.793842436, (double)ws.Evaluate("STDEVPA(A1:A7)"), Tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("STDEVPA({\"hello\"})"));

        AssertScalarToNumberConversion("STDEVPA", 0.5);
        AssertAnyErrorIsPropagated("STDEVPA");
    }

    [Test]
    [Arguments(@"=SUMIF(A1:A10, 1, A1:A10)", 1)]
    [Arguments(@"=SUMIF(A1:A10, 2.0, A1:A10)", 2)]
    [Arguments(@"=SUMIF(A1:A10, 3, A1:A10)", 3)]
    [Arguments(@"=SUMIF(A1:A10, ""3"", A1:A10)", 3)]
    [Arguments(@"=SUMIF(A1:A10, 43831, A1:A10)", 43831)]
    [Arguments(@"=SUMIF(A1:A10, DATE(2020, 1, 1), A1:A10)", 43831)]
    [Arguments(@"=SUMIF(A1:A10, TRUE, A1:A10)", 0)]
    public void SumIfMixedData(string formula, double expected)
    {
        // We follow to Excel's convention.
        // Excel treats 1 and TRUE as unequal, but 3 and "3" as equal
        // LibreOffice Calc handles some SUMIF and COUNTIF differently, e.g. it treats 1 and TRUE as equal, but 3 and "3" differently
        IXLWorksheet ws = this.workbook.Worksheet("MixedData");
        ClassicAssert.AreEqual(expected, ws.Evaluate(formula));
    }

    [Test]
    public void SumIfSpecificationExamples()
    {
        // Test examples from specification.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 3;
        ws.Cell("B1").Value = 10;
        ws.Cell("C1").Value = 7;
        ws.Cell("D1").Value = 10;

        ClassicAssert.AreEqual(20, ws.Evaluate("SUMIF(A1:D1,\"=10\")"));
        ClassicAssert.AreEqual(27, ws.Evaluate("SUMIF(A1:D1,\">5\")"));
        ClassicAssert.AreEqual(10, ws.Evaluate("SUMIF(A1:D1,\"<>10\")"));

        ws.Cell("A2").Value = "apples";
        ws.Cell("B2").Value = "melons";
        ws.Cell("C2").Value = 10;
        ws.Cell("D2").Value = 15;
        ClassicAssert.AreEqual(10, ws.Evaluate("SUMIF(A2:B2,\"*es\",C2:D2)"));
    }

    [Test]
    [Arguments("COUNT(G:I,G:G,H:I)", 258d, DisplayName = "COUNT overlapping columns")]
    [Arguments("COUNT(6:8,6:6,7:8)", 30d, DisplayName = "COUNT overlapping rows")]
    [Arguments("COUNTBLANK(H:J)", 3145640d, DisplayName = "COUNTBLANK columns")]
    [Arguments("COUNTBLANK(7:9)", 49128d, DisplayName = "COUNTBLANK rows")]
    [Arguments("COUNT(1:1048576)", 216d, DisplayName = "COUNT worksheet")]
    [Arguments("COUNTBLANK(1:1048576)", 17179868831d, DisplayName = "COUNTBLANK worksheet")]
    [Arguments("SUM(H:J)", 20501.15d, DisplayName = "SUM columns")]
    [Arguments("SUM(4:5)", 85366.12d, DisplayName = "SUM rows")]
    [Arguments("SUMIF(G:G,50,H:H)", 24.98d, DisplayName = "SUMIF columns")]
    [Arguments("SUMIF(G23:G52,\"\",H3:H32)", 53.24d, DisplayName = "SUMIF ranges")]
    [Arguments("SUMIFS(H:H,G:G,50,I:I,\">900\")", 19.99d, DisplayName = "SUMIFS columns")]
    public void TallySkipsEmptyCells(string formulaA1, double expectedResult)
    {
        using (XLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();
            //Let's pre-initialize cells we need so they didn't affect the result
            ws.Range("A1:J45").Style.Fill.BackgroundColor = XLColor.Amber;
            ws.Cell("ZZ1000").Value = 1;

            double actualResult = (double)ws.Evaluate(formulaA1);

            ClassicAssert.AreEqual(expectedResult, actualResult, Tolerance);
        }
    }

    [Test]
    public void Var()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        // Example from specification
        ClassicAssert.AreEqual(2683.2, ws.Evaluate("VAR(1202,1220,1323,1254,1302)"));

        // Only non-convertible text in D column, thus less than 2 samples.
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("VAR(D3:D45)"));

        // Calculate VAR from numeric values (reference contains only numbers)
        ClassicAssert.AreEqual(2241.560169, (double)ws.Evaluate("VAR(H3:H45)"), Tolerance);

        // Ignores text values in the H column and only uses numeric ones, same as reference with only number
        ClassicAssert.AreEqual(2241.560169, (double)ws.Evaluate("VAR(H:H)"), Tolerance);
        ClassicAssert.AreEqual(
            2241.560169,
            (double)this.workbook.Evaluate("VAR(Data!H:H)"),
            Tolerance
        );

        // Need at least two samples, otherwise returns error
        ClassicAssert.AreEqual(XLError.DivisionByZero, this.workbook.Evaluate("VAR({\"hello\"})"));
        ClassicAssert.AreEqual(XLError.DivisionByZero, this.workbook.Evaluate("VAR(5)"));
        ClassicAssert.AreEqual(0.5, this.workbook.Evaluate("VAR(5, 6)"));

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(0.5, this.workbook.Evaluate("VAR({0, 1, \"Hello\", FALSE, TRUE})"));

        // Reference argument only uses number, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        ClassicAssert.AreEqual(0.5, ws.Evaluate("VAR(Z1:Z6)"));

        AssertScalarToNumberConversion("VAR", 0.5);
        AssertAnyErrorIsPropagated("VAR");
    }

    [Test]
    public void VarA()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Example from specification
        ClassicAssert.AreEqual(
            2683.2,
            (double)ws.Evaluate("VARA(1202, 1220, 1323, 1254, 1302)"),
            Tolerance
        );

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            2,
            (double)ws.Evaluate("VARA({5, 7, \"9\", \"Hello\", FALSE, TRUE})"),
            Tolerance
        );

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        ClassicAssert.AreEqual(9.366666667, (double)ws.Evaluate("VARA(A1:A7)"), Tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("VARA({\"hello\"})"));

        AssertScalarToNumberConversion("VARA", 0.5);
        AssertAnyErrorIsPropagated("VARA");
    }

    [Test]
    public void VarP()
    {
        IXLWorksheet ws = this.workbook.Worksheets.First();

        // Example from specification
        ClassicAssert.AreEqual(
            2146.56,
            (double)ws.Evaluate("VARP(1202,1220,1323,1254,1302)"),
            Tolerance
        );

        // Only non-convertible text in D column, thus less than 1 sample.
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("VARP(D3:D45)"));

        // Calculate VARP from numeric values (reference contains only numbers)
        ClassicAssert.AreEqual(2189.430863, (double)ws.Evaluate("VARP(H3:H45)"), Tolerance);

        // Ignores text values in the H column and only uses numeric ones, same as reference with only number
        ClassicAssert.AreEqual(2189.430863, (double)ws.Evaluate("VARP(H:H)"), Tolerance);
        ClassicAssert.AreEqual(
            2189.430863,
            (double)this.workbook.Evaluate("VARP(Data!H:H)"),
            Tolerance
        );

        // Need at least one sample, otherwise returns error
        ClassicAssert.AreEqual(XLError.DivisionByZero, this.workbook.Evaluate("VARP({\"hello\"})"));
        ClassicAssert.AreEqual(0, this.workbook.Evaluate("VARP(5)"));

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            0.25,
            this.workbook.Evaluate("VARP({0, 1, \"Hello\", FALSE, TRUE})")
        );

        // Reference argument only uses number, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        ClassicAssert.AreEqual(0.25, ws.Evaluate("VARP(Z1:Z6)"));

        AssertScalarToNumberConversion("VARP", 0.25);
        AssertAnyErrorIsPropagated("VARP");
    }

    [Test]
    public void VarPa()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Example from specification
        ClassicAssert.AreEqual(
            2146.56,
            (double)ws.Evaluate("VARPA(1202, 1220, 1323, 1254, 1302)"),
            Tolerance
        );

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(
            1,
            (double)ws.Evaluate("VARPA({5, 7, \"9\", \"Hello\", FALSE, TRUE})"),
            Tolerance
        );

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        ClassicAssert.AreEqual(7.805555556, (double)ws.Evaluate("VARPA(A1:A7)"), Tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("VARPA({\"hello\"})"));

        AssertScalarToNumberConversion("VARPA", 0.25);
        AssertAnyErrorIsPropagated("VARPA");
    }

    [Test]
    public void Large()
    {
        IXLWorksheet ws = this.workbook.Worksheet("Data");
        XLCellValue value = ws.Evaluate("LARGE(G1:G45, 1)");
        ClassicAssert.AreEqual(96, value);

        value = ws.Evaluate("LARGE(G1:G45, 7)");
        ClassicAssert.AreEqual(87, value);

        value = ws.Evaluate("LARGE(G1:G45, 0)");
        ClassicAssert.AreEqual(XLError.NumberInvalid, value);

        value = ws.Evaluate("LARGE(G1:G45, -1)");
        ClassicAssert.AreEqual(XLError.NumberInvalid, value);

        value = ws.Evaluate("LARGE(G1:G45,\"test\")");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, value);

        value = ws.Evaluate("LARGE(C:C,7)");
        ClassicAssert.AreEqual(42623, value);

        value = ws.Evaluate("LARGE(D:D,7)");
        ClassicAssert.AreEqual(XLError.NumberInvalid, value);

        ws = this.workbook.Worksheet("MixedData");

        value = ws.Evaluate("LARGE(A1:A7,6)");
        ClassicAssert.AreEqual(XLError.NumberInvalid, value);

        // Ignores non-numbers.
        value = ws.Evaluate("LARGE(A1:A7,5)");
        ClassicAssert.AreEqual(1, value);

        // Accepts non-area references.
        value = ws.Evaluate("LARGE((A1:A2,A4:A6),2)");
        ClassicAssert.AreEqual(3, value);

        // Errors are returned.
        value = ws.Evaluate("LARGE({ 1, 2, #N/A }, 1)");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, value);

        // Uses ceiling logic for number (1.1 -> 2) + can use arrays.
        value = ws.Evaluate("LARGE({ 1, 2 }, 1.1)");
        ClassicAssert.AreEqual(1, value);

        // If a scalar number-like value supplied, it is converted to number.
        value = ws.Evaluate("LARGE(\"1 1/2\", 1)");
        ClassicAssert.AreEqual(1.5, value);

        // When the scalar can't be converted, return conversion error.
        value = ws.Evaluate("LARGE(\"test\", 1)");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, value);
    }

    private static XLWorkbook SetupWorkbook()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet("Data");
        object[] data =
        [
            new
            {
                Id = 1,
                OrderDate = DateTime.Parse("2015-01-06"),
                Region = "East",
                Rep = "Jones",
                Item = "Pencil",
                Units = 95,
                UnitCost = 1.99,
                Total = 189.05,
            },
            new
            {
                Id = 2,
                OrderDate = DateTime.Parse("2015-01-23"),
                Region = "Central",
                Rep = "Kivell",
                Item = "Binder",
                Units = 50,
                UnitCost = 19.99,
                Total = 999.5,
            },
            new
            {
                Id = 3,
                OrderDate = DateTime.Parse("2015-02-09"),
                Region = "Central",
                Rep = "Jardine",
                Item = "Pencil",
                Units = 36,
                UnitCost = 4.99,
                Total = 179.64,
            },
            new
            {
                Id = 4,
                OrderDate = DateTime.Parse("2015-02-26"),
                Region = "Central",
                Rep = "Gill",
                Item = "Pen",
                Units = 27,
                UnitCost = 19.99,
                Total = 539.73,
            },
            new
            {
                Id = 5,
                OrderDate = DateTime.Parse("2015-03-15"),
                Region = "West",
                Rep = "Sorvino",
                Item = "Pencil",
                Units = 56,
                UnitCost = 2.99,
                Total = 167.44,
            },
            new
            {
                Id = 6,
                OrderDate = DateTime.Parse("2015-04-01"),
                Region = "East",
                Rep = "Jones",
                Item = "Binder",
                Units = 60,
                UnitCost = 4.99,
                Total = 299.4,
            },
            new
            {
                Id = 7,
                OrderDate = DateTime.Parse("2015-04-18"),
                Region = "Central",
                Rep = "Andrews",
                Item = "Pencil",
                Units = 75,
                UnitCost = 1.99,
                Total = 149.25,
            },
            new
            {
                Id = 8,
                OrderDate = DateTime.Parse("2015-05-05"),
                Region = "Central",
                Rep = "Jardine",
                Item = "Pencil",
                Units = 90,
                UnitCost = 4.99,
                Total = 449.1,
            },
            new
            {
                Id = 9,
                OrderDate = DateTime.Parse("2015-05-22"),
                Region = "West",
                Rep = "Thompson",
                Item = "Pencil",
                Units = 32,
                UnitCost = 1.99,
                Total = 63.68,
            },
            new
            {
                Id = 10,
                OrderDate = DateTime.Parse("2015-06-08"),
                Region = "East",
                Rep = "Jones",
                Item = "Binder",
                Units = 60,
                UnitCost = 8.99,
                Total = 539.4,
            },
            new
            {
                Id = 11,
                OrderDate = DateTime.Parse("2015-06-25"),
                Region = "Central",
                Rep = "Morgan",
                Item = "Pencil",
                Units = 90,
                UnitCost = 4.99,
                Total = 449.1,
            },
            new
            {
                Id = 12,
                OrderDate = DateTime.Parse("2015-07-12"),
                Region = "East",
                Rep = "Howard",
                Item = "Binder",
                Units = 29,
                UnitCost = 1.99,
                Total = 57.71,
            },
            new
            {
                Id = 13,
                OrderDate = DateTime.Parse("2015-07-29"),
                Region = "East",
                Rep = "Parent",
                Item = "Binder",
                Units = 81,
                UnitCost = 19.99,
                Total = 1619.19,
            },
            new
            {
                Id = 14,
                OrderDate = DateTime.Parse("2015-08-15"),
                Region = "East",
                Rep = "Jones",
                Item = "Pencil",
                Units = 35,
                UnitCost = 4.99,
                Total = 174.65,
            },
            new
            {
                Id = 15,
                OrderDate = DateTime.Parse("2015-09-01"),
                Region = "Central",
                Rep = "Smith",
                Item = "Desk",
                Units = 2,
                UnitCost = 125,
                Total = 250,
            },
            new
            {
                Id = 16,
                OrderDate = DateTime.Parse("2015-09-18"),
                Region = "East",
                Rep = "Jones",
                Item = "Pen Set",
                Units = 16,
                UnitCost = 15.99,
                Total = 255.84,
            },
            new
            {
                Id = 17,
                OrderDate = DateTime.Parse("2015-10-05"),
                Region = "Central",
                Rep = "Morgan",
                Item = "Binder",
                Units = 28,
                UnitCost = 8.99,
                Total = 251.72,
            },
            new
            {
                Id = 18,
                OrderDate = DateTime.Parse("2015-10-22"),
                Region = "East",
                Rep = "Jones",
                Item = "Pen",
                Units = 64,
                UnitCost = 8.99,
                Total = 575.36,
            },
            new
            {
                Id = 19,
                OrderDate = DateTime.Parse("2015-11-08"),
                Region = "East",
                Rep = "Parent",
                Item = "Pen",
                Units = 15,
                UnitCost = 19.99,
                Total = 299.85,
            },
            new
            {
                Id = 20,
                OrderDate = DateTime.Parse("2015-11-25"),
                Region = "Central",
                Rep = "Kivell",
                Item = "Pen Set",
                Units = 96,
                UnitCost = 4.99,
                Total = 479.04,
            },
            new
            {
                Id = 21,
                OrderDate = DateTime.Parse("2015-12-12"),
                Region = "Central",
                Rep = "Smith",
                Item = "Pencil",
                Units = 67,
                UnitCost = 1.29,
                Total = 86.43,
            },
            new
            {
                Id = 22,
                OrderDate = DateTime.Parse("2015-12-29"),
                Region = "East",
                Rep = "Parent",
                Item = "Pen Set",
                Units = 74,
                UnitCost = 15.99,
                Total = 1183.26,
            },
            new
            {
                Id = 23,
                OrderDate = DateTime.Parse("2016-01-15"),
                Region = "Central",
                Rep = "Gill",
                Item = "Binder",
                Units = 46,
                UnitCost = 8.99,
                Total = 413.54,
            },
            new
            {
                Id = 24,
                OrderDate = DateTime.Parse("2016-02-01"),
                Region = "Central",
                Rep = "Smith",
                Item = "Binder",
                Units = 87,
                UnitCost = 15,
                Total = 1305,
            },
            new
            {
                Id = 25,
                OrderDate = DateTime.Parse("2016-02-18"),
                Region = "East",
                Rep = "Jones",
                Item = "Binder",
                Units = 4,
                UnitCost = 4.99,
                Total = 19.96,
            },
            new
            {
                Id = 26,
                OrderDate = DateTime.Parse("2016-03-07"),
                Region = "West",
                Rep = "Sorvino",
                Item = "Binder",
                Units = 7,
                UnitCost = 19.99,
                Total = 139.93,
            },
            new
            {
                Id = 27,
                OrderDate = DateTime.Parse("2016-03-24"),
                Region = "Central",
                Rep = "Jardine",
                Item = "Pen Set",
                Units = 50,
                UnitCost = 4.99,
                Total = 249.5,
            },
            new
            {
                Id = 28,
                OrderDate = DateTime.Parse("2016-04-10"),
                Region = "Central",
                Rep = "Andrews",
                Item = "Pencil",
                Units = 66,
                UnitCost = 1.99,
                Total = 131.34,
            },
            new
            {
                Id = 29,
                OrderDate = DateTime.Parse("2016-04-27"),
                Region = "East",
                Rep = "Howard",
                Item = "Pen",
                Units = 96,
                UnitCost = 4.99,
                Total = 479.04,
            },
            new
            {
                Id = 30,
                OrderDate = DateTime.Parse("2016-05-14"),
                Region = "Central",
                Rep = "Gill",
                Item = "Pencil",
                Units = 53,
                UnitCost = 1.29,
                Total = 68.37,
            },
            new
            {
                Id = 31,
                OrderDate = DateTime.Parse("2016-05-31"),
                Region = "Central",
                Rep = "Gill",
                Item = "Binder",
                Units = 80,
                UnitCost = 8.99,
                Total = 719.2,
            },
            new
            {
                Id = 32,
                OrderDate = DateTime.Parse("2016-06-17"),
                Region = "Central",
                Rep = "Kivell",
                Item = "Desk",
                Units = 5,
                UnitCost = 125,
                Total = 625,
            },
            new
            {
                Id = 33,
                OrderDate = DateTime.Parse("2016-07-04"),
                Region = "East",
                Rep = "Jones",
                Item = "Pen Set",
                Units = 62,
                UnitCost = 4.99,
                Total = 309.38,
            },
            new
            {
                Id = 34,
                OrderDate = DateTime.Parse("2016-07-21"),
                Region = "Central",
                Rep = "Morgan",
                Item = "Pen Set",
                Units = 55,
                UnitCost = 12.49,
                Total = 686.95,
            },
            new
            {
                Id = 35,
                OrderDate = DateTime.Parse("2016-08-07"),
                Region = "Central",
                Rep = "Kivell",
                Item = "Pen Set",
                Units = 42,
                UnitCost = 23.95,
                Total = 1005.9,
            },
            new
            {
                Id = 36,
                OrderDate = DateTime.Parse("2016-08-24"),
                Region = "West",
                Rep = "Sorvino",
                Item = "Desk",
                Units = 3,
                UnitCost = 275,
                Total = 825,
            },
            new
            {
                Id = 37,
                OrderDate = DateTime.Parse("2016-09-10"),
                Region = "Central",
                Rep = "Gill",
                Item = "Pencil",
                Units = 7,
                UnitCost = 1.29,
                Total = 9.03,
            },
            new
            {
                Id = 38,
                OrderDate = DateTime.Parse("2016-09-27"),
                Region = "West",
                Rep = "Sorvino",
                Item = "Pen",
                Units = 76,
                UnitCost = 1.99,
                Total = 151.24,
            },
            new
            {
                Id = 39,
                OrderDate = DateTime.Parse("2016-10-14"),
                Region = "West",
                Rep = "Thompson",
                Item = "Binder",
                Units = 57,
                UnitCost = 19.99,
                Total = 1139.43,
            },
            new
            {
                Id = 40,
                OrderDate = DateTime.Parse("2016-10-31"),
                Region = "Central",
                Rep = "Andrews",
                Item = "Pencil",
                Units = 14,
                UnitCost = 1.29,
                Total = 18.06,
            },
            new
            {
                Id = 41,
                OrderDate = DateTime.Parse("2016-11-17"),
                Region = "Central",
                Rep = "Jardine",
                Item = "Binder",
                Units = 11,
                UnitCost = 4.99,
                Total = 54.89,
            },
            new
            {
                Id = 42,
                OrderDate = DateTime.Parse("2016-12-04"),
                Region = "Central",
                Rep = "Jardine",
                Item = "Binder",
                Units = 94,
                UnitCost = 19.99,
                Total = 1879.06,
            },
            new
            {
                Id = 43,
                OrderDate = DateTime.Parse("2016-12-21"),
                Region = "Central",
                Rep = "Andrews",
                Item = "Binder",
                Units = 28,
                UnitCost = 4.99,
                Total = 139.72,
            },
        ];

        ws1.FirstCell().CellBelow().CellRight().InsertTable(data, "Table1");

        IXLWorksheet ws2 = wb.AddWorksheet("MixedData");
        ws2.FirstCell()
            .InsertData(
                new object[]
                {
                    1,
                    2.0,
                    "3",
                    3,
                    new DateTime(2020, 1, 1),
                    true,
                    new TimeSpan(10, 5, 30, 10),
                }
            );

        return wb;
    }

    private static void AssertScalarToNumberConversion(string functionName, double result)
    {
        // Scalar blank is converted to 0
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"{functionName}(IF(TRUE,), 1)"),
            Tolerance
        );

        // Scalar logical is converted to a number
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"{functionName}(FALSE, TRUE)"),
            Tolerance
        );
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"{functionName}(0, TRUE)"),
            Tolerance
        );
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"{functionName}(FALSE, 1)"),
            Tolerance
        );

        // Scalar text is converted to a number
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"{functionName}(\"0\", \"1\")"),
            Tolerance
        );
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"{functionName}(\"1\", \"0 0/2\")"),
            Tolerance
        );

        // Scalar text that is not convertible returns error
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"{functionName}(5, \"Hello\")")
        );
    }

    /// <summary>
    /// Assert that a function propagates any error, whether from scalar, array or reference argument.
    /// </summary>
    /// <param name="functionName">Name of a function that accepts any value as argument.</param>
    private static void AssertAnyErrorIsPropagated(string functionName)
    {
        // Scalar error is propagated
        ClassicAssert.AreEqual(
            XLError.NullValue,
            XLWorkbook.EvaluateExpr($"{functionName}(1, #NULL!)")
        );

        // Array error is propagated
        ClassicAssert.AreEqual(
            XLError.NullValue,
            XLWorkbook.EvaluateExpr($"{functionName}({{1, #NULL!}})")
        );

        // Reference error is propagated
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B1").Value = XLError.NoValueAvailable;
        ws.Cell("B2").Value = 1;
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate($"{functionName}(B1)"));
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate($"{functionName}(B1:B2)"));
    }
}
