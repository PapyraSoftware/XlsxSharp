using System;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.CalcEngine;

[TestFixture]
[SetCulture("en-US")]
public class InformationTests
{
    [TestCase("A1")] // blank
    [TestCase("TRUE")]
    [TestCase("14.5")]
    [TestCase("\"text\"")]
    public void ErrorTypeNonErrorsAreNa(string argumentFormula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate($"ERROR.TYPE({argumentFormula})"));
    }

    [TestCase("#NULL!", 1)]
    [TestCase("#DIV/0!", 2)]
    [TestCase("#VALUE!", 3)]
    [TestCase("#REF!", 4)]
    [TestCase("#NAME?", 5)]
    [TestCase("#NUM!", 6)]
    [TestCase("#N/A", 7)]
    //[TestCase("#GETTING_DATA", 8)] OLAP Cube not supported
    public void ErrorTypeReturnsNumberForError(string error, int expectedNumber) =>
        Assert.AreEqual(expectedNumber, XLWorkbook.EvaluateExpr($"ERROR.TYPE({error})"));

    #region IsBlank Tests

    [Test]
    public void IsBlankEmptyCellTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate("IsBlank(A1)");
        Assert.AreEqual(true, actual);
    }

    [Test]
    public void IsBlankNonEmptyCellFalse()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "1";
        XLCellValue actual = ws.Evaluate("IsBlank(A1)");
        Assert.AreEqual(false, actual);
    }

    [TestCase("FALSE")]
    [TestCase("0")]
    [TestCase("5")]
    [TestCase("\"\"")]
    [TestCase("\"Hello\"")]
    [TestCase("#DIV/0!")]
    public void IsBlankNonEmptyValueFalse(string value)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsBlank({value})");
        Assert.AreEqual(false, actual);
    }

    [Test]
    public void IsBlankInlineBlankTrue()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("IsBlank(IF(TRUE,,))");
        Assert.AreEqual(true, actual);
    }

    #endregion IsBlank Tests

    [TestCase("IF(TRUE,,)")]
    [TestCase("FALSE")]
    [TestCase("0")]
    [TestCase("\"\"")]
    [TestCase("\"text\"")]
    public void IsErrNonErrorValuesFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
        Assert.AreEqual(false, actual);
    }

    [TestCase("#DIV/0!")]
    [TestCase("#NAME?")]
    [TestCase("#NULL!")]
    [TestCase("#NUM!")]
    [TestCase("#REF!")]
    [TestCase("#VALUE!")]
    public void IsErrErrorsExceptNaTrue(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
        Assert.AreEqual(true, actual);
    }

    [Test]
    public void IsErrNaFalse()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("IsErr(#N/A)");
        Assert.AreEqual(false, actual);
    }

    [TestCase("#DIV/0!")]
    [TestCase("#N/A")]
    [TestCase("#NAME?")]
    [TestCase("#NULL!")]
    [TestCase("#NUM!")]
    [TestCase("#REF!")]
    [TestCase("#VALUE!")]
    public void IsErrorErrorsTrue(string error)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsError({error})");
        Assert.AreEqual(true, actual);
    }

    [TestCase("IF(TRUE,,)")]
    [TestCase("FALSE")]
    [TestCase("0")]
    [TestCase("\"\"")]
    [TestCase("\"text\"")]
    public void IsErrorNonErrorsFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsError({valueFormula})");
        Assert.AreEqual(false, actual);
    }

    #region IsEven Tests

    [TestCase("2")]
    [TestCase("\"1 2/2\"")]
    [TestCase("\"4 1/2\"")]
    [TestCase("\"48:30:00\"")]
    [TestCase("\"1900-01-02\"")]
    public void IsEvenNumberLikeValueConvertedThroughValueSemantic(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsEven({valueFormula})");
        Assert.AreEqual(true, actual);
    }

    [Test]
    public void IsEvenNonIntegerValuesTruncatedForEvaluation()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");

        ws.Cell("A1").Value = 4;
        ws.Cell("A2").Value = 0.9;
        ws.Cell("A3").Value = -2.9;

        XLCellValue actual = ws.Evaluate("=IsEven(A1)");
        Assert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsEven(A2)");
        Assert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsEven(A3)");
        Assert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsEven(A4)");
        Assert.AreEqual(true, actual);
    }

    [Test]
    [Ignore("Arrays not yet implemented.")]
    public void IsEvenArrayReturnsArray() =>
        Assert.AreEqual(2.0, XLWorkbook.EvaluateExpr("SUM(N(IsEven({\"2.9\";2;1})))"));

    [Test]
    public void IsEvenReferenceToMoreThanOneCellError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 2).FormulaA1 = "IsEven(A1:A2)";
        Assert.AreEqual(XLError.IncompatibleValue, ws.Cell(1, 2).Value);
    }

    [TestCase("TRUE", XLError.IncompatibleValue)]
    [TestCase("FALSE", XLError.IncompatibleValue)]
    [TestCase("\"\"", XLError.IncompatibleValue)]
    [TestCase("\"test\"", XLError.IncompatibleValue)]
    [TestCase("#DIV/0!", XLError.DivisionByZero)]
    [TestCase("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
    public void IsEvenNonNumberValuesError(string valueFormula, XLError expectedError) =>
        Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr($"IsEven({valueFormula})"));

    #endregion IsEven Tests

    #region IsLogical Tests

    [TestCase("TRUE")]
    [TestCase("FALSE")]
    public void IsLogicalOnlyLogicalTrue(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
        Assert.AreEqual(true, actual);
    }

    [TestCase("IF(TRUE,,)")]
    [TestCase("0")]
    [TestCase("1")]
    [TestCase("\"\"")]
    [TestCase("\"text\"")]
    [TestCase("#NAME?")]
    [TestCase("#N/A")]
    [TestCase("#VALUE!")]
    [TestCase("#REF!")]
    public void IsLogicalNonLogicalValueFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
        Assert.AreEqual(false, actual);
    }

    [Test]
    public void IsLogicalReferenceToLogicalValueTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = true;

        XLCellValue actual = ws.Evaluate("IsLogical(A1)");
        Assert.AreEqual(true, actual);
    }

    #endregion IsLogical Tests

    [Test]
    public void IsNanaTrue()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("ISNA(#N/A)");
        Assert.AreEqual(true, actual);
    }

    [TestCase("IF(TRUE,,)")]
    [TestCase("TRUE")]
    [TestCase("0")]
    [TestCase("\"\"")]
    [TestCase("#REF!")]
    [TestCase("\"#N/A\"")]
    public void IsNaNonNotAvailableValueFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"ISNA({valueFormula})");
        Assert.AreEqual(false, actual);
    }

    #region IsNotText Tests

    [Test]
    public void IsNotTextReferenceToBlankCellTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate("IsNonText(A1)");
        Assert.AreEqual(true, actual);
    }

    [TestCase("")]
    [TestCase("  ")]
    [TestCase("text")]
    public void IsNotTextReferenceToStringCellFalse(string text)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = text;
        XLCellValue actual = ws.Evaluate("IsNonText(A1)");
        Assert.AreEqual(false, actual);
    }

    [Test]
    public void IsNotTextNonTextValuesTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = 123; //Double Value
        ws.Cell("A2").Value = DateTime.Now; //Date Value
        ws.Cell("A3").Value = true; //Bool Value
        ws.Cell("A4").Value = XLError.IncompatibleValue; //Error value

        XLCellValue actual = ws.Evaluate("IsNonText(A1)");
        Assert.AreEqual(true, actual);
        actual = ws.Evaluate("IsNonText(A2)");
        Assert.AreEqual(true, actual);
        actual = ws.Evaluate("IsNonText(A3)");
        Assert.AreEqual(true, actual);
        actual = ws.Evaluate("IsNonText(A4)");
        Assert.AreEqual(true, actual);
    }

    #endregion IsNotText Tests

    #region IsNumber Tests

    [Test]
    public void IsNumberSimpleFalse()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = "asd"; //String Value
        ws.Cell("A2").Value = true; //Bool Value

        XLCellValue actual = ws.Evaluate("IsNumber(A1)");
        Assert.AreEqual(false, actual);
        actual = ws.Evaluate("IsNumber(A2)");
        Assert.AreEqual(false, actual);
    }

    [Test]
    public void IsNumberSimpleTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = 123; //Double Value
        ws.Cell("A2").Value = DateTime.Now; //Date Value
        ws.Cell("A3").Value = new TimeSpan(2, 30, 50); //TimeSpan Value

        XLCellValue actual = ws.Evaluate("=IsNumber(A1)");
        Assert.AreEqual(true, actual);
        actual = ws.Evaluate("=IsNumber(A2)");
        Assert.AreEqual(true, actual);
        actual = ws.Evaluate("=IsNumber(A3)");
        Assert.AreEqual(true, actual);
    }

    [TestCase("TRUE")]
    [TestCase("FALSE")]
    [TestCase("\"\"")]
    [TestCase("#DIV/0!")]
    [TestCase("#NULL!")]
    [TestCase("#VALUE!")]
    [TestCase("#N/A")]
    public void IsNumberNonNumberFalse(string nonNumberValue)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsNumber({nonNumberValue})");
        Assert.AreEqual(false, actual);
    }

    #endregion IsNumber Tests

    #region IsOdd Test

    [SetCulture("en-US")]
    [TestCase("1")]
    [TestCase("\"2 3/3\"")]
    [TestCase("\"5 1/3\"")]
    [TestCase("\"25:30:00\"")]
    [TestCase("\"1900-01-03\"")]
    public void IsOddSingleValueConvertedThroughValueSemantic(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})");
        Assert.AreEqual(true, actual);
    }

    [Test]
    public void IsOddNonIntegerValuesTruncatedForEvaluation()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");

        ws.Cell("A1").Value = 3;
        ws.Cell("A2").Value = 1.9;
        ws.Cell("A3").Value = -5.9;

        XLCellValue actual = ws.Evaluate("=IsOdd(A1)");
        Assert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsOdd(A2)");
        Assert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsOdd(A3)");
        Assert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsOdd(A4)");
        Assert.AreEqual(false, actual);
    }

    [SetCulture("en-US")]
    [Test]
    [Ignore("Arrays not yet implemented.")]
    public void IsOddArrayReturnsArray() =>
        Assert.AreEqual(2.0, XLWorkbook.EvaluateExpr("SUM(N(IsOdd({\"3.2\",7,2})))"));

    [Test]
    public void IsOddReferenceToMoreThanOneCellError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 2).FormulaA1 = "IsOdd(A1:A2)";
        Assert.AreEqual(XLError.IncompatibleValue, ws.Cell(1, 2).Value);
    }

    [TestCase("TRUE", XLError.IncompatibleValue)]
    [TestCase("FALSE", XLError.IncompatibleValue)]
    [TestCase("\"\"", XLError.IncompatibleValue)]
    [TestCase("\"test\"", XLError.IncompatibleValue)]
    [TestCase("#DIV/0!", XLError.DivisionByZero)]
    [TestCase("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
    public void IsOddNonNumberValuesError(string valueFormula, XLError expectedError) =>
        Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})"));

    #endregion IsOdd Test

    [TestCase("A1")]
    [TestCase("(A1,A5)")]
    public void IsRefReferenceTrue(string reference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = "123";

        ws.Cell("B1").FormulaA1 = $"ISREF({reference})";

        Assert.AreEqual(true, ws.Cell("B1").Value);
    }

    [TestCase("IF(TRUE,,)")]
    [TestCase("TRUE")]
    [TestCase("0")]
    [TestCase("\"\"")]
    // [TestCase("{1;2}")] Arrays not yet implemented
    [TestCase("#N/A")]
    [TestCase("#VALUE!")]
    public void IsRefNonReferenceFalse(string nonReference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");

        ws.Cell("B1").FormulaA1 = $"ISREF({nonReference})";

        Assert.AreEqual(false, ws.Cell("B1").Value);
    }

    #region IsText Tests

    [Test]
    public void IsTextBlankCellFalse()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B1").FormulaA1 = "ISTEXT(A1)";

        Assert.AreEqual(false, ws.Cell("B1").Value);
    }

    [TestCase("0")]
    [TestCase("123")]
    [TestCase("TRUE")]
    [TestCase("#DIV/0!")]
    [TestCase("IF(TRUE,,)")]
    public void IsTextNonTextFalse(string nonText)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"ISTEXT({nonText})");
        Assert.AreEqual(false, actual);
    }

    [TestCase("")]
    [TestCase("abc")]
    public void IsTextCellWithTextTrue(string textValue)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = textValue;

        XLCellValue actual = ws.Evaluate("IsText(A1)");
        Assert.AreEqual(true, actual);
    }

    #endregion IsText Tests

    #region N Tests

    [Test]
    public void NBlankZero()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate("N(A1)");
        Assert.AreEqual(0.0, actual);
    }

    [Test]
    public void NDateSerialNumber()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        DateTime testedDate = DateTime.Now;
        ws.Cell("A1").Value = testedDate;
        XLCellValue actual = ws.Evaluate("N(A1)");
        Assert.AreEqual(testedDate.ToSerialDateTime(), actual);
    }

    [Test]
    public void NFalseZero()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = false;
        XLCellValue actual = ws.Evaluate("N(A1)");
        Assert.AreEqual(0, actual);
    }

    [Test]
    public void NTrueOne()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = true;
        XLCellValue actual = ws.Evaluate("N(A1)");
        Assert.AreEqual(1, actual);
    }

    [Test]
    public void NNumberNumber()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        int testedValue = 123;
        ws.Cell("A1").Value = testedValue;
        XLCellValue actual = ws.Evaluate("N(A1)");
        Assert.AreEqual(testedValue, actual);
    }

    [TestCase("")]
    [TestCase("abc")]
    public void NStringZero(string text)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = text;
        XLCellValue actual = ws.Evaluate("N(A1)");
        Assert.AreEqual(0, actual);
    }

    [Test]
    [Ignore("Array not implemented")]
    public void NArrayConvertsIndividualItems()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("SUM(N({2,TRUE}))");
        Assert.AreEqual(3, actual);
    }

    [TestCase("A1")]
    [TestCase("A1:B1")]
    [TestCase("(A1, B1)")]
    public void NReferenceTakesFirstCellFromFirstArea(string reference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        ws.Cell("B1").Value = 10;

        XLCellValue actual = ws.Evaluate($"SUM(N({reference}))");
        Assert.AreEqual(5, actual);
    }

    #endregion N Tests

    [TestCase("IF(TRUE,,)", 1)]
    [TestCase("0", 1)]
    [TestCase("1", 1)]
    [TestCase("-5.2", 1)]
    [TestCase("\"\"", 2)]
    [TestCase("\"text\"", 2)]
    [TestCase("\"1\"", 2)]
    [TestCase("\"TRUE\"", 2)]
    [TestCase("TRUE", 4)]
    [TestCase("FALSE", 4)]
    [TestCase("#DIV/0!", 16)]
    [TestCase("1/0", 16)]
    [TestCase("#N/A", 16)]
    [TestCase("#VALUE!", 16)]
    public void TypeNonReferenceScalarValues(string literalValues, double expectedNumber)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = $"TYPE({literalValues})";
        Assert.AreEqual(expectedNumber, ws.Cell("A1").Value);
    }

    [Ignore("Arrays not implemented")]
    [TestCase("{1}")]
    [TestCase("{TRUE,#N/A}")]
    [TestCase("{\"abc\";5}")]
    public void TypeArrayHasValue64(string arrayLiteral)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"TYPE({arrayLiteral})");
        Assert.AreEqual(64.0, actual);
    }

    [TestCase("A1:A2")]
    // [TestCase("(A1:A3 A2:B3)")] Not implemented // Intersection results in a 1x2 block
    public void TypeReferenceToNonSingleCellBehavesLikeArray(string reference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = $"TYPE({reference})";
        Assert.AreEqual(64.0, ws.Cell("C1").Value);
    }

    [Test]
    public void TypeReferenceToSingleCellReturnsTypeOfCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";

        ws.Cell("C1").FormulaA1 = "TYPE(A1)";
        Assert.AreEqual(2.0, ws.Cell("C1").Value);
    }

    [Test]
    public void TypeMultiAreaReferenceReturnsError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";

        ws.Cell("C1").FormulaA1 = "TYPE((A1,A1))";
        Assert.AreEqual(16.0, ws.Cell("C1").Value);
    }
}
