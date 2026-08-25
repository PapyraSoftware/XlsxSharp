using System.Globalization;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using Array = XlsxSharp.Excel.CalcEngine.Array;

namespace XlsxSharp.Tests.Excel.CalcEngine;

/// <summary>
/// Tests that verify that we can parse formulas and evaluate them. Take a look at XLParser ExcelFormulaGrammar.cs and each rule + its transformation into Abstract Syntax Tree is checked here.
/// </summary>
public class FormulaParserTests
{
    #region Start.Rule

    [Test]
    [Arguments]
    public void FormulaStringCanStartingWithAnEqualSign() =>
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("=1"));

    [Test]
    [Arguments]
    public void FormulaStringCanOmitStartingEqualSign() =>
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("1"));

    [Test]
    [Arguments]
    public void RootFormulaStringCanBeUnionWithoutParenthesis()
    {
        // Root of a formula string is pretty much the only place where reference union can be without parenthesis. Elsewhere it must have
        // parentheses to avoid misusing union op (coma) with a separation of arguments in a function call.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Evaluate("=A1,A3", "Z100");
    }

    #endregion

    #region Formula.Rule

    [Test]
    [Arguments]
    public void FormulaCanBeReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Text";
        ClassicAssert.AreEqual("Text", ws.Evaluate("=A1"));
    }

    [Test]
    [Arguments("=1", 1)]
    [Arguments("=\"text\"", "text")]
    [Arguments("=TRUE", true)]
    public void FormulaCanBeConstant(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=SUM(1,2)", 3)]
    [Arguments("=2+3", 5)]
    [Arguments("=-3", -3)]
    [Arguments("=150%", 1.5)]
    public void FormulaCanBeFunctionCall(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments]
    public void FormulaCanBeConstantArray() =>
        // 1 is determined through implicit intersection (first element)
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("={1,2,3;4,5,6}"));

    [Test]
    [Arguments("=(1)", 1)]
    [Arguments("=(\"text\")", "text")]
    public void FormulaCanBeAnotherFormulaInParenthesis(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    #endregion

    #region Constant.Rule
    [Test]
    [Arguments("=1", 1)] // int
    [Arguments("=1.5", 1.5)] // double
    [Arguments("=1.23e2", 123)]
    [Arguments("=1.23e-1", 0.123)]
    [Arguments("=1.23e+3", 1230)]
    [Arguments("=032399977109", 32399977109)] // long
    [Arguments("=9223372036854775808", 9223372036854775808)] // BigInteger (long value + 1)
    public void ConstantCanBeNumber(string formula, double expectedNumber) =>
        // Irony returns number as an object of various types, e.g. int or double
        ClassicAssert.AreEqual(expectedNumber, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=\"text\"", "text")]
    [Arguments("=\"first line\nsecond line\"", "first line\nsecond line")]
    [Arguments("=\"we'll\"", "we'll")]
    [Arguments(
        "=\"use two double quote \"\" to nest quotes\"",
        "use two double quote \" to nest quotes"
    )]
    public void ConstantCanBeText(string formula, string expectedText) =>
        ClassicAssert.AreEqual(expectedText, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=TRUE", true)]
    [Arguments("=FALSE", false)]
    [Arguments("=tRuE", true)]
    public void ConstantCanBeBool(string formula, bool expectedBool) =>
        ClassicAssert.AreEqual(expectedBool, XLWorkbook.EvaluateExpr(formula));

    // #REF! is converted by a different rule, so it is not here.
    [Test]
    [Arguments("#VALUE!", XLError.IncompatibleValue)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("#NAME?", XLError.NameNotRecognized)]
    [Arguments("#N/A", XLError.NoValueAvailable)]
    [Arguments("#NULL!", XLError.NullValue)]
    [Arguments("#NUM!", XLError.NumberInvalid)]
    public void ConstantCanBeError(string formula, object expectedError)
    {
        XLError error = (XLError)XLWorkbook.EvaluateExpr(formula);
        ClassicAssert.AreEqual(expectedError, error);
    }
    #endregion

    // Function call from XLParser is anything that takes arguments and uses some transformation (e.g. addition, excel function, unary operation..)
    #region FunctionCall.Rule

    [Test]
    [Arguments("=COS(0)", 1)]
    [Arguments("=SUM(1,2,3)", 6)]
    public void FunctionCallCanBeExcelPredefinedFunction(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=+1", 1)]
    [Arguments("=-1", -1)]
    //        [TestCase("=@A1", 1)]
    public void FunctionCallCanBeUnaryPrefixOperation(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=75%", 0.75)]
    public void FunctionCallCanBeUnaryPostfixOperation(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=2^3", 8)]
    [Arguments("=4^1.5", 8)]
    [Arguments("=3*2", 6)]
    [Arguments("=6/2", 3)]
    [Arguments("=3/2", 1.5)]
    [Arguments("=1+2", 3)]
    [Arguments("=3-5", -2)]
    [Arguments(@"=""A"" & ""B""", "AB")]
    [Arguments("=2>1", true)]
    [Arguments("=1>2", false)]
    [Arguments("=5=5", true)]
    [Arguments("=1=2", false)]
    [Arguments("=1<2", true)]
    [Arguments("=2<1", false)]
    [Arguments("=2<>1", true)]
    [Arguments("=3<>3", false)]
    [Arguments("=2>=1", true)]
    [Arguments("=2>=2", true)]
    [Arguments("=1>=2", false)]
    [Arguments("=1<=2", true)]
    [Arguments("=1<=1", true)]
    [Arguments("=2<=1", false)]
    public void FunctionCallCanBeBinaryInfixOperation(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    #endregion

    #region Argument.Rule

    [Test]
    [Arguments("=PMT(0,1,1000,,1)", -1000)]
    public void EmptyArgumentsArePassedToFunction(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    #endregion

    #region Reference.Rule

    [Test]
    [Arguments("=A1", 1)]
    [Arguments("=TestRangeName", 5)]
    //        [TestCase("=UndefinedRangeName", Error.NameNotRecognized)]
    public void ReferenceCanBeReferenceItem(string formula, object expectedValue)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 5;
        ws.Range("A2:A2").AddToNamed("TestRangeName");

        ClassicAssert.AreEqual(expectedValue, ws.Evaluate(formula));
    }

    [Test]
    [Arguments]
    public void ReferenceCanBeReferenceFunctionCall() =>
        // XLParser considers a limited subset of predefined functions (IF, CHOOSE, INDEX...) to be different from other predefined function because they can return reference.
        ClassicAssert.AreEqual(2, XLWorkbook.EvaluateExpr("=IF(FALSE,1,2)"));

    [Test]
    [Arguments]
    public void ReferenceCanBeAnotherReferenceInParenthesis()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        ClassicAssert.AreEqual(1, ws.Evaluate("=(A1)"));
    }

    [Test]
    [Arguments]
    public void ReferenceCanBeReferenceItemWithPrefix()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
        IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");
        ws2.Cell("A1").Value = 1;

        ClassicAssert.AreEqual(1, ws1.Evaluate("=Sheet2!  A1"));
    }

    [Test]
    [Arguments]
    [Skip("XLParser issue #57")]
    public void ReferenceCanBeDynamicDataExchange() =>
        AssertCanParseButNotEvaluate(
            "=Sdemo123|tik!'id1?req?AAPL_STK_SMART_USD_~/'",
            "Evaluation of dynamic data exchange is not implemented."
        );

    #endregion

    #region ReferenceFunctionCall.Rule

    [Test]
    [Arguments]
    public void ReferenceFunctionCallCanBeBinaryRangeOfTwoReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Evaluate("A1:A3:C2", "Z100");
    }

    [Test]
    [Arguments]
    public void ReferenceFunctionCallCanBeIntersectionOfTwoReferences() =>
        AssertCanParseButNotEvaluate(
            "=A1:A3 A2:B2",
            "Evaluation of range intersection operator is not implemented."
        );

    [Test]
    [Arguments]
    public void ReferenceFunctionCallCanBeUnionInParenthesis()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Evaluate("=(A1:A3,A2:B2,B1:B4)", "Z100");
    }

    [Test]
    [Arguments]
    public void ReferenceFunctionCallCanBeReferenceFunction() =>
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("=IF(TRUE,1,2)"));

    [Test]
    [Arguments]
    public void ReferenceFunctionCallCanBeReferenceWithSpillRangeOperator() =>
        AssertCanParseButNotEvaluate(
            "=A1#",
            "Evaluation of spill range operator is not implemented."
        );

    #endregion

    #region RefFunctionName.Rule

    [Test]
    [Arguments("=IF(FALSE,1,2)", 2)]
    // [TestCase("=CHOOSE(2,\"A\",\"B\",73)", "B")] Not implemented
    public void RefFunctionNameCanBeExcelRefConditionalFunction(
        string formula,
        object expectedValue
    ) => ClassicAssert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("=INDEX(A1:B2,1,2)", "Lemons")]
    //[TestCase("=OFFSET(C4,-1,-2)", "Pears")] Not implemented
    //[TestCase("=INDIRECT(\"A2\")", "Bananas")] Not implemented
    public void RefFunctionNameCanBeExcelRefFunction(string formula, object expectedValue)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Apples";
        ws.Cell("B1").Value = "Lemons";
        ws.Cell("A2").Value = "Bananas";
        ws.Cell("B2").Value = "Pears";
        ClassicAssert.AreEqual(expectedValue, ws.Evaluate(formula));
    }

    #endregion

    #region ReferenceItem.Rule
    // Reference item is transient and is thus inside the reference

    [Test]
    [Arguments]
    public void ReferenceItemCanBeCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        ClassicAssert.AreEqual(1, ws.Evaluate("=A1"));
    }

    [Test]
    [Arguments("TestRange")]
    [Arguments("A1A1")]
    public void ReferenceItemCanBeNamedRange(string rangeName)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1).AddToNamed(rangeName);

        ClassicAssert.AreEqual(12, ws.Evaluate($"=SUM({rangeName})"));
    }

    [Test]
    [Arguments]
    public void ReferenceItemCanBeVerticalRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1);

        ClassicAssert.AreEqual(8, ws.Evaluate("=SUM(A:B)"));
    }

    [Test]
    [Arguments]
    public void ReferenceItemCanBeHorizontalRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1);

        ClassicAssert.AreEqual(3, ws.Evaluate("=SUM(2:2)"));
    }

    [Test]
    [Arguments]
    public void ReferenceItemCanBeRefError() =>
        ClassicAssert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("#REF!"));

    [Test]
    [Arguments]
    public void ReferenceItemCanBeUserDefinedFunctionCall() =>
        ClassicAssert.AreEqual(
            XLError.NameNotRecognized,
            XLWorkbook.EvaluateExpr("CustomFunction(1)")
        );

    [Test]
    [Arguments]
    public void ReferenceItemCanBeStructuredReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertTable(new[] { new { Amount = 1 }, new { Amount = 2 } });

        ClassicAssert.AreEqual(3, ws.Evaluate("SUM(Table1[#Data])"));
    }

    #endregion

    #region ConstantArray.Rule

    [Test]
    public void ConstArrayMustHaveSameNumberOfColumns()
    {
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);
        ExpressionParseException ex = ClassicAssert.Throws<ExpressionParseException>(() =>
            calcEngine.Parse("{1;2,3}")
        )!;
        StringAssert.Contains("Rows of an array don't have same size.", ex.Message);
    }

    [Test]
    public void ConstArrayCantContainImplicitIntersectionOperator()
    {
        // Array literal elements are always literal scalar constants - never a reference, a
        // function call, a nested array, or (as here) the implicit intersection operator.
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);
        ExpressionParseException ex = ClassicAssert.Throws<ExpressionParseException>(() =>
            calcEngine.Parse("{@1}")
        )!;
        StringAssert.Contains("Unable to parse value starting from position 1.", ex.Message);
    }

    [Test]
    [MethodDataSource(nameof(ArrayCases))]
    public void ConstArrayCanHaveOnlyScalars(string formula, object expected)
    {
        ConstArray expectedArray = (ConstArray)expected;
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);

        Formula ast = calcEngine.Parse(formula);

        Array actual = ((ArrayNode)ast.AstRoot).Value;
        ClassicAssert.AreEqual(expectedArray.Width, actual.Width);
        ClassicAssert.AreEqual(expectedArray.Height, actual.Height);
        for (int row = 0; row < actual.Height; ++row)
        {
            for (int col = 0; col < actual.Width; ++col)
            {
                ScalarValue actualElement = actual[row, col];
                ScalarValue expectedElement = expectedArray[row, col];
                ClassicAssert.AreEqual(expectedElement, actualElement);
            }
        }
    }

    internal static IEnumerable<object[]> ArrayCases
    {
        get
        {
            yield return
            [
                "{1}",
                new ConstArray(
                    new ScalarValue[,]
                    {
                        { 1 },
                    }
                ),
            ];
            yield return
            [
                "{#REF!}",
                new ConstArray(
                    new ScalarValue[,]
                    {
                        { XLError.CellReference },
                    }
                ),
            ];
            yield return
            [
                "{1,2,3,4}",
                new ConstArray(
                    new ScalarValue[,]
                    {
                        { 1, 2, 3, 4 },
                    }
                ),
            ];
            yield return
            [
                "{1,2;3,4}",
                new ConstArray(
                    new ScalarValue[,]
                    {
                        { 1, 2 },
                        { 3, 4 },
                    }
                ),
            ];
            yield return
            [
                "{+1,#REF!,\"Text\";FALSE,#DIV/0!,-1.5}",
                new ConstArray(
                    new ScalarValue[,]
                    {
                        { 1, XLError.CellReference, "Text" },
                        { false, XLError.DivisionByZero, -1.5 },
                    }
                ),
            ];
        }
    }

    #endregion

    #region Prefix.Rule

    // No quotes
    [Test]
    [Arguments("=Sheet5!A1", "Sheet5")]
    [Arguments("=Test_sheet!A1", "Test_sheet")]
    // Sheet with quotes
    [Arguments("='Test Sheet'!A1", "Test Sheet")]
    [Arguments("='Test-Sheet'!A1", "Test-Sheet")]
    [Arguments("='^%>;-+'!A1", "^%>;-+")]
    // Sheet can be named as #REF! error, but sheet reference must be escaped
    [Arguments("='#REF'!A1", "#REF")]
    public void PrefixCanBeSheetToken(string formula, string sheetName)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet(sheetName);
        ws.Cell("A1").Value = 5;
        ClassicAssert.AreEqual(5, ws.Evaluate(formula));
    }

    [Test]
    [Arguments("=Sheet1:Sheet5!A1")]
    [Arguments("=Jan:Dec!A1")]
    public void PrefixCanBeSheetsFor3DReference(string formula) =>
        AssertCanParseButNotEvaluate(formula, "3D references are not yet implemented.");

    [Test]
    [Arguments("=[1]Sheet4!A1")]
    public void PrefixCanBeFileAndSheetToken(string formula) =>
        AssertCanParseButNotEvaluate(
            formula,
            "References from other files are not yet implemented."
        );

    #endregion

    private static void AssertCanParseButNotEvaluate(string formula, string notSupportedMessage)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);
        _ = calcEngine.Parse(formula);
        NotImplementedException ex = ClassicAssert.Throws<NotImplementedException>(() =>
            ws.Evaluate(formula, "A1")
        );
        ClassicAssert.AreEqual(notSupportedMessage, ex.Message);
    }
}
