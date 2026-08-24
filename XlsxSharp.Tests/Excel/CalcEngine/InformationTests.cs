using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class InformationTests
{
    [Test]
    [Arguments("A1")] // blank
    [Arguments("TRUE")]
    [Arguments("14.5")]
    [Arguments("\"text\"")]
    public void ErrorTypeNonErrorsAreNa(string argumentFormula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(
            XLError.NoValueAvailable,
            ws.Evaluate($"ERROR.TYPE({argumentFormula})")
        );
    }

    [Test]
    [Arguments("#NULL!", 1)]
    [Arguments("#DIV/0!", 2)]
    [Arguments("#VALUE!", 3)]
    [Arguments("#REF!", 4)]
    [Arguments("#NAME?", 5)]
    [Arguments("#NUM!", 6)]
    [Arguments("#N/A", 7)]
    //[TestCase("#GETTING_DATA", 8)] OLAP Cube not supported
    public void ErrorTypeReturnsNumberForError(string error, int expectedNumber) =>
        ClassicAssert.AreEqual(expectedNumber, XLWorkbook.EvaluateExpr($"ERROR.TYPE({error})"));

    #region IsBlank Tests

    [Test]
    public void IsBlankEmptyCellTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate("IsBlank(A1)");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    public void IsBlankNonEmptyCellFalse()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "1";
        XLCellValue actual = ws.Evaluate("IsBlank(A1)");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    [Arguments("FALSE")]
    [Arguments("0")]
    [Arguments("5")]
    [Arguments("\"\"")]
    [Arguments("\"Hello\"")]
    [Arguments("#DIV/0!")]
    public void IsBlankNonEmptyValueFalse(string value)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsBlank({value})");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    public void IsBlankInlineBlankTrue()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("IsBlank(IF(TRUE,,))");
        ClassicAssert.AreEqual(true, actual);
    }

    #endregion IsBlank Tests

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("FALSE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    [Arguments("\"text\"")]
    public void IsErrNonErrorValuesFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    [Arguments("#DIV/0!")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#REF!")]
    [Arguments("#VALUE!")]
    public void IsErrErrorsExceptNaTrue(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    public void IsErrNaFalse()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("IsErr(#N/A)");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    [Arguments("#DIV/0!")]
    [Arguments("#N/A")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#REF!")]
    [Arguments("#VALUE!")]
    public void IsErrorErrorsTrue(string error)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsError({error})");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("FALSE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    [Arguments("\"text\"")]
    public void IsErrorNonErrorsFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsError({valueFormula})");
        ClassicAssert.AreEqual(false, actual);
    }

    #region IsEven Tests

    [Test]
    [Arguments("2")]
    [Arguments("\"1 2/2\"")]
    [Arguments("\"4 1/2\"")]
    [Arguments("\"48:30:00\"")]
    [Arguments("\"1900-01-02\"")]
    public void IsEvenNumberLikeValueConvertedThroughValueSemantic(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsEven({valueFormula})");
        ClassicAssert.AreEqual(true, actual);
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
        ClassicAssert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsEven(A2)");
        ClassicAssert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsEven(A3)");
        ClassicAssert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsEven(A4)");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    [Skip("Arrays not yet implemented.")]
    public void IsEvenArrayReturnsArray() =>
        ClassicAssert.AreEqual(2.0, XLWorkbook.EvaluateExpr("SUM(N(IsEven({\"2.9\";2;1})))"));

    [Test]
    public void IsEvenReferenceToMoreThanOneCellError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 2).FormulaA1 = "IsEven(A1:A2)";
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell(1, 2).Value);
    }

    [Test]
    [Arguments("TRUE", XLError.IncompatibleValue)]
    [Arguments("FALSE", XLError.IncompatibleValue)]
    [Arguments("\"\"", XLError.IncompatibleValue)]
    [Arguments("\"test\"", XLError.IncompatibleValue)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
    public void IsEvenNonNumberValuesError(string valueFormula, XLError expectedError) =>
        ClassicAssert.AreEqual(expectedError, XLWorkbook.EvaluateExpr($"IsEven({valueFormula})"));

    #endregion IsEven Tests

    #region IsLogical Tests

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    public void IsLogicalOnlyLogicalTrue(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("0")]
    [Arguments("1")]
    [Arguments("\"\"")]
    [Arguments("\"text\"")]
    [Arguments("#NAME?")]
    [Arguments("#N/A")]
    [Arguments("#VALUE!")]
    [Arguments("#REF!")]
    public void IsLogicalNonLogicalValueFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    public void IsLogicalReferenceToLogicalValueTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = true;

        XLCellValue actual = ws.Evaluate("IsLogical(A1)");
        ClassicAssert.AreEqual(true, actual);
    }

    #endregion IsLogical Tests

    [Test]
    public void IsNanaTrue()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("ISNA(#N/A)");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("TRUE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    [Arguments("#REF!")]
    [Arguments("\"#N/A\"")]
    public void IsNaNonNotAvailableValueFalse(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"ISNA({valueFormula})");
        ClassicAssert.AreEqual(false, actual);
    }

    #region IsNotText Tests

    [Test]
    public void IsNotTextReferenceToBlankCellTrue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate("IsNonText(A1)");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    [Arguments("")]
    [Arguments("  ")]
    [Arguments("text")]
    public void IsNotTextReferenceToStringCellFalse(string text)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = text;
        XLCellValue actual = ws.Evaluate("IsNonText(A1)");
        ClassicAssert.AreEqual(false, actual);
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
        ClassicAssert.AreEqual(true, actual);
        actual = ws.Evaluate("IsNonText(A2)");
        ClassicAssert.AreEqual(true, actual);
        actual = ws.Evaluate("IsNonText(A3)");
        ClassicAssert.AreEqual(true, actual);
        actual = ws.Evaluate("IsNonText(A4)");
        ClassicAssert.AreEqual(true, actual);
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
        ClassicAssert.AreEqual(false, actual);
        actual = ws.Evaluate("IsNumber(A2)");
        ClassicAssert.AreEqual(false, actual);
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
        ClassicAssert.AreEqual(true, actual);
        actual = ws.Evaluate("=IsNumber(A2)");
        ClassicAssert.AreEqual(true, actual);
        actual = ws.Evaluate("=IsNumber(A3)");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("\"\"")]
    [Arguments("#DIV/0!")]
    [Arguments("#NULL!")]
    [Arguments("#VALUE!")]
    [Arguments("#N/A")]
    public void IsNumberNonNumberFalse(string nonNumberValue)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsNumber({nonNumberValue})");
        ClassicAssert.AreEqual(false, actual);
    }

    #endregion IsNumber Tests

    #region IsOdd Test

    [Test]
    [Arguments("1")]
    [Arguments("\"2 3/3\"")]
    [Arguments("\"5 1/3\"")]
    [Arguments("\"25:30:00\"")]
    [Arguments("\"1900-01-03\"")]
    public void IsOddSingleValueConvertedThroughValueSemantic(string valueFormula)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})");
        ClassicAssert.AreEqual(true, actual);
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
        ClassicAssert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsOdd(A2)");
        ClassicAssert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsOdd(A3)");
        ClassicAssert.AreEqual(true, actual);

        actual = ws.Evaluate("=IsOdd(A4)");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    [Skip("Arrays not yet implemented.")]
    public void IsOddArrayReturnsArray() =>
        ClassicAssert.AreEqual(2.0, XLWorkbook.EvaluateExpr("SUM(N(IsOdd({\"3.2\",7,2})))"));

    [Test]
    public void IsOddReferenceToMoreThanOneCellError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 2).FormulaA1 = "IsOdd(A1:A2)";
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell(1, 2).Value);
    }

    [Test]
    [Arguments("TRUE", XLError.IncompatibleValue)]
    [Arguments("FALSE", XLError.IncompatibleValue)]
    [Arguments("\"\"", XLError.IncompatibleValue)]
    [Arguments("\"test\"", XLError.IncompatibleValue)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
    public void IsOddNonNumberValuesError(string valueFormula, XLError expectedError) =>
        ClassicAssert.AreEqual(expectedError, XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})"));

    #endregion IsOdd Test

    [Test]
    [Arguments("A1")]
    [Arguments("(A1,A5)")]
    public void IsRefReferenceTrue(string reference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = "123";

        ws.Cell("B1").FormulaA1 = $"ISREF({reference})";

        ClassicAssert.AreEqual(true, ws.Cell("B1").Value);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("TRUE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    // [TestCase("{1;2}")] Arrays not yet implemented
    [Arguments("#N/A")]
    [Arguments("#VALUE!")]
    public void IsRefNonReferenceFalse(string nonReference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");

        ws.Cell("B1").FormulaA1 = $"ISREF({nonReference})";

        ClassicAssert.AreEqual(false, ws.Cell("B1").Value);
    }

    #region IsText Tests

    [Test]
    public void IsTextBlankCellFalse()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B1").FormulaA1 = "ISTEXT(A1)";

        ClassicAssert.AreEqual(false, ws.Cell("B1").Value);
    }

    [Test]
    [Arguments("0")]
    [Arguments("123")]
    [Arguments("TRUE")]
    [Arguments("#DIV/0!")]
    [Arguments("IF(TRUE,,)")]
    public void IsTextNonTextFalse(string nonText)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"ISTEXT({nonText})");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    [Arguments("")]
    [Arguments("abc")]
    public void IsTextCellWithTextTrue(string textValue)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = textValue;

        XLCellValue actual = ws.Evaluate("IsText(A1)");
        ClassicAssert.AreEqual(true, actual);
    }

    #endregion IsText Tests

    #region N Tests

    [Test]
    public void NBlankZero()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate("N(A1)");
        ClassicAssert.AreEqual(0.0, actual);
    }

    [Test]
    public void NDateSerialNumber()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        DateTime testedDate = DateTime.Now;
        ws.Cell("A1").Value = testedDate;
        XLCellValue actual = ws.Evaluate("N(A1)");
        ClassicAssert.AreEqual(testedDate.ToSerialDateTime(), actual);
    }

    [Test]
    public void NFalseZero()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = false;
        XLCellValue actual = ws.Evaluate("N(A1)");
        ClassicAssert.AreEqual(0, actual);
    }

    [Test]
    public void NTrueOne()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = true;
        XLCellValue actual = ws.Evaluate("N(A1)");
        ClassicAssert.AreEqual(1, actual);
    }

    [Test]
    public void NNumberNumber()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        int testedValue = 123;
        ws.Cell("A1").Value = testedValue;
        XLCellValue actual = ws.Evaluate("N(A1)");
        ClassicAssert.AreEqual(testedValue, actual);
    }

    [Test]
    [Arguments("")]
    [Arguments("abc")]
    public void NStringZero(string text)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = text;
        XLCellValue actual = ws.Evaluate("N(A1)");
        ClassicAssert.AreEqual(0, actual);
    }

    [Test]
    [Skip("Array not implemented")]
    public void NArrayConvertsIndividualItems()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("SUM(N({2,TRUE}))");
        ClassicAssert.AreEqual(3, actual);
    }

    [Test]
    [Arguments("A1")]
    [Arguments("A1:B1")]
    [Arguments("(A1, B1)")]
    public void NReferenceTakesFirstCellFromFirstArea(string reference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        ws.Cell("B1").Value = 10;

        XLCellValue actual = ws.Evaluate($"SUM(N({reference}))");
        ClassicAssert.AreEqual(5, actual);
    }

    #endregion N Tests

    [Test]
    [Arguments("IF(TRUE,,)", 1)]
    [Arguments("0", 1)]
    [Arguments("1", 1)]
    [Arguments("-5.2", 1)]
    [Arguments("\"\"", 2)]
    [Arguments("\"text\"", 2)]
    [Arguments("\"1\"", 2)]
    [Arguments("\"TRUE\"", 2)]
    [Arguments("TRUE", 4)]
    [Arguments("FALSE", 4)]
    [Arguments("#DIV/0!", 16)]
    [Arguments("1/0", 16)]
    [Arguments("#N/A", 16)]
    [Arguments("#VALUE!", 16)]
    public void TypeNonReferenceScalarValues(string literalValues, double expectedNumber)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = $"TYPE({literalValues})";
        ClassicAssert.AreEqual(expectedNumber, ws.Cell("A1").Value);
    }

    [Test]
    [Skip("Arrays not implemented")]
    [Arguments("{1}")]
    [Arguments("{TRUE,#N/A}")]
    [Arguments("{\"abc\";5}")]
    public void TypeArrayHasValue64(string arrayLiteral)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"TYPE({arrayLiteral})");
        ClassicAssert.AreEqual(64.0, actual);
    }

    [Test]
    [Arguments("A1:A2")]
    // [TestCase("(A1:A3 A2:B3)")] Not implemented // Intersection results in a 1x2 block
    public void TypeReferenceToNonSingleCellBehavesLikeArray(string reference)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = $"TYPE({reference})";
        ClassicAssert.AreEqual(64.0, ws.Cell("C1").Value);
    }

    [Test]
    public void TypeReferenceToSingleCellReturnsTypeOfCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";

        ws.Cell("C1").FormulaA1 = "TYPE(A1)";
        ClassicAssert.AreEqual(2.0, ws.Cell("C1").Value);
    }

    [Test]
    public void TypeMultiAreaReferenceReturnsError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";

        ws.Cell("C1").FormulaA1 = "TYPE((A1,A1))";
        ClassicAssert.AreEqual(16.0, ws.Cell("C1").Value);
    }
}
