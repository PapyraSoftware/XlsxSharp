using System;
using System.Collections.Generic;
using System.Globalization;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using Array = XlsxSharp.Excel.CalcEngine.Array;

namespace XlsxSharp.Tests.Excel.CalcEngine;

/// <summary>
/// Tests that verify that we can parse formulas and evaluate them. Take a look at XLParser ExcelFormulaGrammar.cs and each rule + its transformation into Abstract Syntax Tree is checked here.
/// </summary>
[TestFixture]
public class FormulaParserTests
{
    #region Start.Rule

    [TestCase]
    public void FormulaStringCanStartingWithAnEqualSign() =>
        Assert.AreEqual(1, XLWorkbook.EvaluateExpr("=1"));

    [TestCase]
    public void FormulaStringCanOmitStartingEqualSign() =>
        Assert.AreEqual(1, XLWorkbook.EvaluateExpr("1"));

    [TestCase]
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

    [TestCase]
    public void FormulaCanBeReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Text";
        Assert.AreEqual("Text", ws.Evaluate("=A1"));
    }

    [TestCase("=1", 1)]
    [TestCase("=\"text\"", "text")]
    [TestCase("=TRUE", true)]
    public void FormulaCanBeConstant(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=SUM(1,2)", 3)]
    [TestCase("=2+3", 5)]
    [TestCase("=-3", -3)]
    [TestCase("=150%", 1.5)]
    public void FormulaCanBeFunctionCall(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [TestCase]
    public void FormulaCanBeConstantArray() =>
        // 1 is determined through implicit intersection (first element)
        Assert.AreEqual(1, XLWorkbook.EvaluateExpr("={1,2,3;4,5,6}"));

    [TestCase("=(1)", 1)]
    [TestCase("=(\"text\")", "text")]
    public void FormulaCanBeAnotherFormulaInParenthesis(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    #endregion

    #region Constant.Rule
    [TestCase("=1", 1)] // int
    [TestCase("=1.5", 1.5)] // double
    [TestCase("=1.23e2", 123)]
    [TestCase("=1.23e-1", 0.123)]
    [TestCase("=1.23e+3", 1230)]
    [TestCase("=032399977109", 32399977109)] // long
    [TestCase("=9223372036854775808", 9223372036854775808)] // BigInteger (long value + 1)
    public void ConstantCanBeNumber(string formula, double expectedNumber) =>
        // Irony returns number as an object of various types, e.g. int or double
        Assert.AreEqual(expectedNumber, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=\"text\"", "text")]
    [TestCase("=\"first line\nsecond line\"", "first line\nsecond line")]
    [TestCase("=\"we'll\"", "we'll")]
    [TestCase(
        "=\"use two double quote \"\" to nest quotes\"",
        "use two double quote \" to nest quotes"
    )]
    public void ConstantCanBeText(string formula, string expectedText) =>
        Assert.AreEqual(expectedText, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=TRUE", true)]
    [TestCase("=FALSE", false)]
    [TestCase("=tRuE", true)]
    public void ConstantCanBeBool(string formula, bool expectedBool) =>
        Assert.AreEqual(expectedBool, XLWorkbook.EvaluateExpr(formula));

    // #REF! is converted by a different rule, so it is not here.
    [TestCase("#VALUE!", XLError.IncompatibleValue)]
    [TestCase("#DIV/0!", XLError.DivisionByZero)]
    [TestCase("#NAME?", XLError.NameNotRecognized)]
    [TestCase("#N/A", XLError.NoValueAvailable)]
    [TestCase("#NULL!", XLError.NullValue)]
    [TestCase("#NUM!", XLError.NumberInvalid)]
    public void ConstantCanBeError(string formula, object expectedError)
    {
        XLError error = (XLError)XLWorkbook.EvaluateExpr(formula);
        Assert.AreEqual(expectedError, error);
    }
    #endregion

    // Function call from XLParser is anything that takes arguments and uses some transformation (e.g. addition, excel function, unary operation..)
    #region FunctionCall.Rule

    [TestCase("=COS(0)", 1)]
    [TestCase("=SUM(1,2,3)", 6)]
    public void FunctionCallCanBeExcelPredefinedFunction(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=+1", 1)]
    [TestCase("=-1", -1)]
    //        [TestCase("=@A1", 1)]
    public void FunctionCallCanBeUnaryPrefixOperation(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=75%", 0.75)]
    public void FunctionCallCanBeUnaryPostfixOperation(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=2^3", 8)]
    [TestCase("=4^1.5", 8)]
    [TestCase("=3*2", 6)]
    [TestCase("=6/2", 3)]
    [TestCase("=3/2", 1.5)]
    [TestCase("=1+2", 3)]
    [TestCase("=3-5", -2)]
    [TestCase(@"=""A"" & ""B""", "AB")]
    [TestCase("=2>1", true)]
    [TestCase("=1>2", false)]
    [TestCase("=5=5", true)]
    [TestCase("=1=2", false)]
    [TestCase("=1<2", true)]
    [TestCase("=2<1", false)]
    [TestCase("=2<>1", true)]
    [TestCase("=3<>3", false)]
    [TestCase("=2>=1", true)]
    [TestCase("=2>=2", true)]
    [TestCase("=1>=2", false)]
    [TestCase("=1<=2", true)]
    [TestCase("=1<=1", true)]
    [TestCase("=2<=1", false)]
    public void FunctionCallCanBeBinaryInfixOperation(string formula, object expectedValue) =>
        Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    #endregion

    #region Argument.Rule

    [TestCase("=PMT(0,1,1000,,1)", -1000)]
    public void EmptyArgumentsArePassedToFunction(string formula, object expectedValue) =>
        Assert.That(XLWorkbook.EvaluateExpr(formula), Is.EqualTo(expectedValue));

    #endregion

    #region Reference.Rule

    [TestCase("=A1", 1)]
    [TestCase("=TestRangeName", 5)]
    //        [TestCase("=UndefinedRangeName", Error.NameNotRecognized)]
    public void ReferenceCanBeReferenceItem(string formula, object expectedValue)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 5;
        ws.Range("A2:A2").AddToNamed("TestRangeName");

        Assert.AreEqual(expectedValue, ws.Evaluate(formula));
    }

    [TestCase]
    public void ReferenceCanBeReferenceFunctionCall() =>
        // XLParser considers a limited subset of predefined functions (IF, CHOOSE, INDEX...) to be different from other predefined function because they can return reference.
        Assert.AreEqual(2, XLWorkbook.EvaluateExpr("=IF(FALSE,1,2)"));

    [TestCase]
    public void ReferenceCanBeAnotherReferenceInParenthesis()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        Assert.AreEqual(1, ws.Evaluate("=(A1)"));
    }

    [TestCase]
    public void ReferenceCanBeReferenceItemWithPrefix()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
        IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");
        ws2.Cell("A1").Value = 1;

        Assert.AreEqual(1, ws1.Evaluate("=Sheet2!  A1"));
    }

    [TestCase]
    [Ignore("XLParser issue #57")]
    public void ReferenceCanBeDynamicDataExchange() =>
        AssertCanParseButNotEvaluate(
            "=Sdemo123|tik!'id1?req?AAPL_STK_SMART_USD_~/'",
            "Evaluation of dynamic data exchange is not implemented."
        );

    #endregion

    #region ReferenceFunctionCall.Rule

    [TestCase]
    public void ReferenceFunctionCallCanBeBinaryRangeOfTwoReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Evaluate("A1:A3:C2", "Z100");
    }

    [TestCase]
    public void ReferenceFunctionCallCanBeIntersectionOfTwoReferences() =>
        AssertCanParseButNotEvaluate(
            "=A1:A3 A2:B2",
            "Evaluation of range intersection operator is not implemented."
        );

    [TestCase]
    public void ReferenceFunctionCallCanBeUnionInParenthesis()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Evaluate("=(A1:A3,A2:B2,B1:B4)", "Z100");
    }

    [TestCase]
    public void ReferenceFunctionCallCanBeReferenceFunction() =>
        Assert.AreEqual(1, XLWorkbook.EvaluateExpr("=IF(TRUE,1,2)"));

    [TestCase]
    public void ReferenceFunctionCallCanBeReferenceWithSpillRangeOperator() =>
        AssertCanParseButNotEvaluate(
            "=A1#",
            "Evaluation of spill range operator is not implemented."
        );

    #endregion

    #region RefFunctionName.Rule

    [TestCase("=IF(FALSE,1,2)", 2)]
    // [TestCase("=CHOOSE(2,\"A\",\"B\",73)", "B")] Not implemented
    public void RefFunctionNameCanBeExcelRefConditionalFunction(
        string formula,
        object expectedValue
    ) => Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));

    [TestCase("=INDEX(A1:B2,1,2)", "Lemons")]
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
        Assert.AreEqual(expectedValue, ws.Evaluate(formula));
    }

    #endregion

    #region ReferenceItem.Rule
    // Reference item is transient and is thus inside the reference

    [TestCase]
    public void ReferenceItemCanBeCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        Assert.AreEqual(1, ws.Evaluate("=A1"));
    }

    [TestCase("TestRange")]
    [TestCase("A1A1")]
    public void ReferenceItemCanBeNamedRange(string rangeName)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1).AddToNamed(rangeName);

        Assert.AreEqual(12, ws.Evaluate($"=SUM({rangeName})"));
    }

    [TestCase]
    public void ReferenceItemCanBeVerticalRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1);

        Assert.AreEqual(8, ws.Evaluate("=SUM(A:B)"));
    }

    [TestCase]
    public void ReferenceItemCanBeHorizontalRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1);

        Assert.AreEqual(3, ws.Evaluate("=SUM(2:2)"));
    }

    [TestCase]
    public void ReferenceItemCanBeRefError() =>
        Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("#REF!"));

    [TestCase]
    public void ReferenceItemCanBeUserDefinedFunctionCall() =>
        Assert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("CustomFunction(1)"));

    [TestCase]
    public void ReferenceItemCanBeStructuredReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertTable(new[] { new { Amount = 1 }, new { Amount = 2 } });

        Assert.AreEqual(3, ws.Evaluate("SUM(Table1[#Data])"));
    }

    #endregion

    #region ConstantArray.Rule

    [Test]
    public void ConstArrayMustHaveSameNumberOfColumns()
    {
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);
        ExpressionParseException ex = Assert.Throws<ExpressionParseException>(() =>
            calcEngine.Parse("{1;2,3}")
        )!;
        StringAssert.Contains("Rows of an array don't have same size.", ex.Message);
    }

    [Test]
    public void ConstArrayCantContainImplicitIntersectionOperator()
    {
        // XLParser allows @ for number through 'PrefixOp + Number'
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);
        ExpressionParseException ex = Assert.Throws<ExpressionParseException>(() =>
            calcEngine.Parse("{@1}")
        )!;
        StringAssert.Contains("Unexpected token INTERSECT.", ex.Message);
    }

    [TestCaseSource(nameof(ArrayCases))]
    public void ConstArrayCanHaveOnlyScalars(string formula, object expected)
    {
        ConstArray expectedArray = (ConstArray)expected;
        XLCalcEngine calcEngine = new(CultureInfo.InvariantCulture);

        Formula ast = calcEngine.Parse(formula);

        Array actual = ((ArrayNode)ast.AstRoot).Value;
        Assert.AreEqual(expectedArray.Width, actual.Width);
        Assert.AreEqual(expectedArray.Height, actual.Height);
        for (int row = 0; row < actual.Height; ++row)
        {
            for (int col = 0; col < actual.Width; ++col)
            {
                ScalarValue actualElement = actual[row, col];
                ScalarValue expectedElement = expectedArray[row, col];
                Assert.AreEqual(expectedElement, actualElement);
            }
        }
    }

    private static IEnumerable<object[]> ArrayCases
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
    [TestCase("=Sheet5!A1", "Sheet5")]
    [TestCase("=Test_sheet!A1", "Test_sheet")]
    // Sheet with quotes
    [TestCase("='Test Sheet'!A1", "Test Sheet")]
    [TestCase("='Test-Sheet'!A1", "Test-Sheet")]
    [TestCase("='^%>;-+'!A1", "^%>;-+")]
    // Sheet can be named as #REF! error, but sheet reference must be escaped
    [TestCase("='#REF'!A1", "#REF")]
    public void PrefixCanBeSheetToken(string formula, string sheetName)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet(sheetName);
        ws.Cell("A1").Value = 5;
        Assert.AreEqual(5, ws.Evaluate(formula));
    }

    [TestCase("=Sheet1:Sheet5!A1")]
    [TestCase("=Jan:Dec!A1")]
    public void PrefixCanBeSheetsFor3DReference(string formula) =>
        AssertCanParseButNotEvaluate(formula, "3D references are not yet implemented.");

    [TestCase("=[1]Sheet4!A1")]
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
        Assert.Throws(
            Is.TypeOf<NotImplementedException>().With.Message.EqualTo(notSupportedMessage),
            () => ws.Evaluate(formula, "A1")
        );
    }
}
