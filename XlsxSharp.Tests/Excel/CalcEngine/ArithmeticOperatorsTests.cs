using System;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class ArithmeticOperatorsTests
{
    #region Concat text operator

    [Test]
    [Arguments("\"A\" & \"B\"", "AB")]
    [Arguments("\"\" & \"B\"", "B")]
    [Arguments("\"A\" & \"\"", "A")]
    [Arguments("\"\" & \"\"", "")]
    public void ConcatConcatenateText(string formula, object expectedResult) =>
        ClassicAssert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("A1 & \"\"", "")]
    [Arguments("\"\" & A1", "")]
    [Arguments("A1 & A1", "")]
    public void ConcatConcatenateBlank(string formula, object expectedResult) =>
        ClassicAssert.AreEqual(expectedResult, Evaluate(formula));

    [Test]
    [Arguments("TRUE & \" to text\"", "TRUE to text")]
    [Arguments("FALSE & \" to text\"", "FALSE to text")]
    [Arguments("true & \" to text\"", "TRUE to text")]
    [Arguments("false & \" to text\"", "FALSE to text")]
    [Arguments("TRUE & FALSE", @"TRUEFALSE")]
    public void ConcatConvertsLogicalToString(string formula, object expectedResult) =>
        ClassicAssert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Culture("cs-CZ")]
    [Arguments("1 & \" to text\"", "1 to text")]
    [Arguments("1 & 0", "10")]
    [Arguments("1.5 & 0.78", "1,50,78")]
    public void ConcatConvertsNumberToStringUsingCulture(string formula, object expectedResult)
    {
        XLWorkbook wb = new();
        ClassicAssert.AreEqual(expectedResult, wb.Evaluate(formula));
    }

    [Test]
    [Arguments("#DIV/0! & 1", XLError.DivisionByZero)]
    [Arguments("#DIV/0! & \"1\"", XLError.DivisionByZero)]
    [Arguments("#REF! & #DIV/0!", XLError.CellReference)]
    [Arguments("1 & #NAME?", XLError.NameNotRecognized)]
    public void ConcatWithErrorAsOperandReturnsTheError(string formula, XLError expectedError) =>
        ClassicAssert.AreEqual(expectedError, XLWorkbook.EvaluateExpr(formula));

    #endregion

    #region Unary plus

    [Test]
    [Arguments("+1", 1)]
    [Arguments("+\"1\"", "1")]
    [Arguments("+TRUE", true)]
    [Arguments("+FALSE", false)]
    [Arguments("+#DIV/0!", XLError.DivisionByZero)]
    [Arguments("ISBLANK(+A1)", true)]
    public void UnaryPlusIsNonOpThatKeepsValueAndType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Unary minus

    [Test]
    [Arguments("-1", -1)]
    [Arguments("-125.45", -125.45)]
    [Arguments("-\"1\"", -1)]
    [Arguments("-TRUE", -1)]
    [Arguments("-FALSE", 0)]
    [Arguments("-#DIV/0!", XLError.DivisionByZero)]
    [Arguments("-A1", 0.0)]
    public void UnaryMinusConvertsArgumentBeforeNegating(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Unary minus

    [Test]
    [Arguments("1%", 0.01)]
    [Arguments("100%", 1.0)]
    [Arguments("25.7%", 0.257)]
    [Arguments("125.45%", 1.2545)]
    [Arguments("\"1\"%", 0.01)]
    [Arguments("TRUE%", 0.01)]
    [Arguments("FALSE%", 0)]
    [Arguments("#NAME?%", XLError.NameNotRecognized)]
    [Arguments("(1/0)%", XLError.DivisionByZero)]
    [Arguments("A1%", 0.0)]
    public void UnaryPercentConvertsArgumentBeforePercentOperator(
        string formula,
        object expectedValue
    ) => ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Exponentiation

    [Test]
    [Arguments("1^1", 1.0)]
    [Arguments("0^0", XLError.NumberInvalid)]
    [Arguments("10^0", 1.0)]
    [Arguments("4^0.5", 2.0)]
    [Arguments("2^0.5", 1.4142135623730951)]
    [Arguments("2^-2", 0.25)]
    [Arguments("\"5\"^\"3\"", 125)]
    [Arguments("5^TRUE", 5)]
    [Arguments("5^FALSE", 1)]
    [Arguments("#VALUE!^1", XLError.IncompatibleValue)]
    [Arguments("1^#REF!", XLError.CellReference)]
    [Arguments("#DIV/0!^#REF!", XLError.DivisionByZero)]
    [Arguments("5^A1", 1.0)]
    [Arguments("A1^4", 0.0)]
    public void ExponentiationCanWorkWithScalars(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Multiplication

    [Test]
    [Arguments("1+1", 2.0)]
    [Arguments("0*0", 0.0)]
    [Arguments("10*0", 0.0)]
    [Arguments("2*1.5", 3.0)]
    [Arguments("2.5*2.5", 6.25)]
    [Arguments("2*-2", -4)]
    [Arguments("\"5\" * \"3\"", 15)]
    [Arguments("5*TRUE", 5)]
    [Arguments("5*FALSE", 0)]
    [Arguments("#VALUE!*1", XLError.IncompatibleValue)]
    [Arguments("1*#REF!", XLError.CellReference)]
    [Arguments("#DIV/0!*#REF!", XLError.DivisionByZero)]
    [Arguments("10*A1", 0.0)]
    [Arguments("A1*10", 0.0)]
    public void MultiplicationCanWorkWithScalars(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Division

    [Test]
    [Arguments("1/1", 1.0)]
    [Arguments("5/2", 2.5)]
    [Arguments("14.5/2.5", 5.8)]
    [Arguments("10/0", XLError.DivisionByZero)]
    [Arguments("0/0", XLError.DivisionByZero)]
    [Arguments("2.5/-0.5", -5)]
    [Arguments("\"10\" / \"4\"", 2.5)]
    [Arguments("5/TRUE", 5)]
    [Arguments("5/FALSE", XLError.DivisionByZero)]
    [Arguments("#VALUE!/1", XLError.IncompatibleValue)]
    [Arguments("1/#REF!", XLError.CellReference)]
    [Arguments("#DIV/0!/#REF!", XLError.DivisionByZero)]
    [Arguments("A1/5", 0.0)]
    [Arguments("5/A1", XLError.DivisionByZero)]
    public void DivisionCanWorkWithScalars(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Addition

    [Test]
    [Arguments("1+1", 2.0)]
    [Arguments("5+2.5", 7.5)]
    [Arguments("10+0", 10.0)]
    [Arguments("\"10\" + \"4\"", 14.0)]
    [Arguments("5+TRUE", 6.0)]
    [Arguments("5+FALSE", 5.0)]
    [Arguments("#VALUE! + 1", XLError.IncompatibleValue)]
    [Arguments("1 + #REF!", XLError.CellReference)]
    [Arguments("#DIV/0! + #REF!", XLError.DivisionByZero)]
    [Arguments("A1 + 7", 7)]
    public void AdditionCanWorkWithScalars(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Subtraction

    [Test]
    [Arguments("1-1", 0.0)]
    [Arguments("2.5-7.8", -5.3)]
    [Arguments("10-0", 10.0)]
    [Arguments("\"10\" - \"4\"", 6.0)]
    [Arguments("5-TRUE", 4.0)]
    [Arguments("5-FALSE", 5.0)]
    [Arguments("#VALUE! - 1", XLError.IncompatibleValue)]
    [Arguments("1 - #REF!", XLError.CellReference)]
    [Arguments("#DIV/0! - #REF!", XLError.DivisionByZero)]
    [Arguments("A1 - 5", -5)]
    public void SubtractionCanWorkWithScalars(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    #endregion

    #region Array Operations

    [Test]
    public void ArraysOperationBinaryOperationBetweenAreaReferenceAndSingleCellReferenceShouldWork()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Test1");
        ws.Cell("A1").Value = new DateTime(2021, 1, 15);
        ws.Cell("A2").Value = new DateTime(2021, 1, 10);
        ws.Cell("B1").Value = new DateTime(2021, 1, 5);
        ClassicAssert.AreEqual(5, ws.Evaluate("MIN(A1:A2-B1)"));
    }

    [Test]
    public void ArraysOperationMultiAreaReferencesArgumentResultsInScalarError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cells("A1:A2").Value = 1;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("(A1:A1,A1:A2)+1"));
        ClassicAssert.AreEqual(16, ws.Evaluate("TYPE((A1:A1,A1:A2)+1)")); // The result is a scalar error, not an array of errors
    }

    [Test]
    public void ArrayOperationSameSizeArrayPerformsOperationIndividually()
    {
        ClassicAssert.AreEqual(
            6 * 7,
            XLWorkbook.EvaluateExpr("SUM({1,2,3;4,5,6} + {6,5,4;3,2,1})")
        );
        ClassicAssert.AreEqual(2, XLWorkbook.EvaluateExpr("COLUMNS({1,2} + \"A\")"));
    }

    [Test]
    public void ArrayOperationArrayPlusScalarUpscalesScalarToSizeOfArray()
    {
        ClassicAssert.AreEqual(18, XLWorkbook.EvaluateExpr("SUM({1,1,1;1,1,1} * 3)"));
        ClassicAssert.AreEqual(15, XLWorkbook.EvaluateExpr("SUM(6 / {2,2,2;3,3,3})"));
    }

    [Test]
    public void ArrayOperationRowOnlyArrayIsRepeatedToHaveSameNumberOfRowsAsOtherArray()
    {
        // {3,2} is scaled to {3,2;3,2} of second array
        ClassicAssert.AreEqual(14, XLWorkbook.EvaluateExpr("SUM({3,2}+{1,1;1,1})"));
        ClassicAssert.AreEqual(14, XLWorkbook.EvaluateExpr("SUM({1,1;1,1}+{3,2})"));
    }

    [Test]
    public void ArrayOperationColumnOnlyArrayIsRepeatedToHaveSameNumberOfColumnsAsOtherArray()
    {
        // {3;2} is scaled to {3,3;2,2} of second array
        ClassicAssert.AreEqual(16, XLWorkbook.EvaluateExpr("SUM({3;2}*{1,1;2,3})"));
        ClassicAssert.AreEqual(16, XLWorkbook.EvaluateExpr("SUM({1,1;2,3}*{3;2})"));
    }

    [Test]
    public void ArrayOperation1X1ArrayIsScaledToOtherArray()
    {
        ClassicAssert.AreEqual(20, XLWorkbook.EvaluateExpr("SUM({2}*{1,2;3,4})"));
        ClassicAssert.AreEqual(20, XLWorkbook.EvaluateExpr("SUM({1,2;3,4}*{2})"));
    }

    [Test]
    public void ArrayOperationDifferentSizedArraysAreUpscaledToContainingSize()
    {
        // The extra value are #N/A + value, i.e. #N/A, thus the whole sum is #N/A
        ClassicAssert.AreEqual(
            XLError.NoValueAvailable,
            XLWorkbook.EvaluateExpr("SUM({1,2;3,4;5,6}+{1,2,3;4,5,6})")
        );
        ClassicAssert.AreEqual(3, XLWorkbook.EvaluateExpr("ROWS({1,2;3,4;5,6}+{1,2,3;4,5,6})"));
        ClassicAssert.AreEqual(3, XLWorkbook.EvaluateExpr("COLUMNS({1,2;3,4;5,6}+{1,2,3;4,5,6})"));
    }

    #endregion

    private static XLCellValue Evaluate(string formula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        return ws.Evaluate(formula);
    }
}
