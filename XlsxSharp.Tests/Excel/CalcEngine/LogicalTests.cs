using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class LogicalTests
{
    [Test]
    public void AndIsLogicalConjunction()
    {
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND(TRUE)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND(TRUE, TRUE)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND(TRUE, TRUE, TRUE)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND({TRUE, TRUE}, TRUE)"));

        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND(FALSE)"));
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND(TRUE, FALSE)"));
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND({TRUE, FALSE})"));
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND(TRUE, {TRUE, FALSE})"));
    }

    [Test]
    [Arguments("A1")]
    [Arguments("A1:A5")]
    [Arguments("(A1:A5,B1:B5)")]
    public void AndNoCollectionValuesError(string range)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate($"AND({range})"));
    }

    [Test]
    public void AndScalarArgumentsCoercedFromBlankOrTextOrNumber()
    {
        // Blank evaluated to false
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND(IF(TRUE,,))"));

        // Number coerced to logical
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND(0)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND(0.1)"));

        // Text coerced to logical
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("AND(\"FALSE\")"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND(\"TRUE\")"));
    }

    [Test]
    public void AndUnconvertableScalarArgumentsSkipped() =>
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("AND(TRUE,\"z\")"));

    [Test]
    public void AndOnlyLogicalOrNumberElementsOfCollectionUsed()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // 0 is a number and is converted to logical
        ws.Cell("A1").Value = 0;
        ClassicAssert.AreEqual(false, ws.Evaluate("AND(TRUE,A1)"));

        // false is a logical
        ws.Cell("A2").Value = false;
        ClassicAssert.AreEqual(false, ws.Evaluate("AND(TRUE,A2)"));

        // Text is not converted and thus skipped for evaluation
        ws.Cell("A3").Value = "FALSE";
        ClassicAssert.AreEqual(true, ws.Evaluate("AND(TRUE,A3)"));

        ws.Cell("A4").Value = "some text";
        ClassicAssert.AreEqual(true, ws.Evaluate("AND(TRUE,A4)"));
    }

    [Test]
    public void If2ParamsTrue()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"")");
        ClassicAssert.AreEqual("T", actual);
    }

    [Test]
    public void If2ParamsFalse()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"")");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    public void If3ParamsTrue()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"", ""F"")");
        ClassicAssert.AreEqual("T", actual);
    }

    [Test]
    public void If3ParamsFalse()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"", ""F"")");
        ClassicAssert.AreEqual("F", actual);
    }

    [Test]
    public void IfComparingAgainstEmptyString()
    {
        object actual;
        actual = XLWorkbook.EvaluateExpr(@"if(date(2016, 1, 1) = """", ""A"",""B"")");
        ClassicAssert.AreEqual("B", actual);

        actual = XLWorkbook.EvaluateExpr(@"if("""" = date(2016, 1, 1), ""A"",""B"")");
        ClassicAssert.AreEqual("B", actual);

        actual = XLWorkbook.EvaluateExpr(@"if("""" = 123, ""A"",""B"")");
        ClassicAssert.AreEqual("B", actual);

        actual = XLWorkbook.EvaluateExpr(@"if("""" = """", ""A"",""B"")");
        ClassicAssert.AreEqual("A", actual);
    }

    [Test]
    public void IfCaseInsensitivity()
    {
        object actual;
        actual = XLWorkbook.EvaluateExpr(@"IF(""text""=""TEXT"", 1, 2)");
        ClassicAssert.AreEqual(1, actual);
    }

    [Test]
    public void IfCanReturnReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(true, ws.Evaluate("ISREF(IF(TRUE, A1))"));
        ClassicAssert.AreEqual(true, ws.Evaluate("ISREF(IF(FALSE,, A1))"));
    }

    [Test]
    public void IfHasScalarConditionAndRangeValues()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
        ws.Cell("B1").InsertData(new[] { 4, 5, 6 });
        ws.Cell("C1").InsertData(new[] { true, false, true });
        for (int row = 1; row <= 4; ++row)
        {
            ws.Cell(row, 4).FormulaA1 = "SUM(IF(C1:C3, A1:A3, B1:B3))";
        }

        // Condition is implicitely intersected, because it's a scalar parameter
        ClassicAssert.AreEqual(6, ws.Cell("D1").Value);
        ClassicAssert.AreEqual(15, ws.Cell("D2").Value);
        ClassicAssert.AreEqual(6, ws.Cell("D3").Value);
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("D4").Value);
    }

    [Test]
    public void IfConditionErrorReturnError() =>
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExpr(@"IF(1/0, ""T"", ""F"")")
        );

    [Test]
    public void IfConditionCoercedToLogical()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual("F", ws.Evaluate(@"IF(A1, ""T"", ""F"")"));

        ClassicAssert.AreEqual("T", ws.Evaluate(@"IF(""TRUE"", ""T"", ""F"")"));
        ClassicAssert.AreEqual("F", ws.Evaluate(@"IF(""FALSE"", ""T"", ""F"")"));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            ws.Evaluate(@"IF(""text"", ""T"", ""F"")")
        );

        ClassicAssert.AreEqual("T", ws.Evaluate(@"IF(1, ""T"", ""F"")"));
        ClassicAssert.AreEqual("F", ws.Evaluate(@"IF(0, ""T"", ""F"")"));
    }

    [Test]
    public void IfMissingValuesReturnBlank()
    {
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr(@"ISBLANK(IF(TRUE,,))"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr(@"ISBLANK(IF(FALSE,,))"));
    }

    [Test]
    public void IfErrorFirstArgumentNonErrorReturnFirstArgument()
    {
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("ISBLANK(IFERROR(IF(TRUE,), 5))"));

        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("IFERROR(FALSE, 5)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("IFERROR(TRUE, 5)"));

        ClassicAssert.AreEqual(0.0, XLWorkbook.EvaluateExpr("IFERROR(0, 5)"));
        ClassicAssert.AreEqual(-2.0, XLWorkbook.EvaluateExpr("IFERROR(-2, 5)"));

        ClassicAssert.AreEqual(string.Empty, XLWorkbook.EvaluateExpr("IFERROR(\"\", 5)"));
        ClassicAssert.AreEqual("text", XLWorkbook.EvaluateExpr("IFERROR(\"text\", 5)"));
    }

    [Test]
    public void IfErrorFirstArgumentErrorReturnSecondArgument()
    {
        ClassicAssert.AreEqual("text", XLWorkbook.EvaluateExpr("IFERROR(1/0, \"text\")"));

        ClassicAssert.AreEqual(
            XLError.NameNotRecognized,
            XLWorkbook.EvaluateExpr("IFERROR(#REF!, #NAME?)")
        );
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("IFERROR(#NULL!, TRUE)"));
        ClassicAssert.AreEqual(
            true,
            XLWorkbook.EvaluateExpr("ISBLANK(IFERROR(#VALUE!,IF(TRUE,)))")
        );
    }

    [Test]
    public void IfErrorReferenceNeverReturned()
    {
        // Unlike IF, IFERROR doesn't return reference
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(false, ws.Evaluate("ISREF(IFERROR(#VALUE!, A1))"));
    }

    [Test]
    [Arguments("TRUE", false)]
    [Arguments("FALSE", true)]
    [Arguments("IF(TRUE,,)", true)] // Blank
    [Arguments("0", true)]
    [Arguments("0.1", false)]
    [Arguments("\"true\"", false)]
    [Arguments("\"false\"", true)]
    [Arguments("1/0", XLError.DivisionByZero)]
    public void Not(string valueFormula, object expectedResult) =>
        ClassicAssert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr($"NOT({valueFormula})"));

    [Test]
    public void OrIsLogicalDisjunction()
    {
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR(TRUE)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR(TRUE, TRUE)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR(TRUE, FALSE, TRUE)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR({FALSE, TRUE}, FALSE)"));

        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR(FALSE)"));
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR(FALSE, FALSE)"));
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR({FALSE, FALSE})"));
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR(FALSE, {FALSE, FALSE})"));
    }

    [Test]
    [Arguments("A1")]
    [Arguments("A1:A5")]
    [Arguments("(A1:A5,B1:B5)")]
    public void OrNoCollectionValuesError(string range)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate($"OR({range})"));
    }

    [Test]
    public void OrScalarArgumentsCoercedFromBlankOrTextOrNumber()
    {
        // Blank evaluated to false
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR(IF(TRUE,,))"));

        // Number coerced to logical
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR(0)"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR(0.1)"));

        // Text coerced to logical
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("OR(\"FALSE\")"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR(\"TRUE\")"));
    }

    [Test]
    public void OrUnconvertableScalarArgumentsSkipped() =>
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("OR(TRUE,\"z\")"));

    [Test]
    public void OrOnlyLogicalOrNumberElementsOfCollectionUsed()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // 1 is a number and is converted to logical
        ws.Cell("A1").Value = 1;
        ClassicAssert.AreEqual(true, ws.Evaluate("OR(FALSE,A1)"));

        // false is a logical
        ws.Cell("A2").Value = true;
        ClassicAssert.AreEqual(true, ws.Evaluate("OR(FALSE,A2)"));

        // Text is not converted and thus skipped for evaluation
        ws.Cell("A3").Value = "TRUE";
        ClassicAssert.AreEqual(false, ws.Evaluate("OR(FALSE,A3)"));

        ws.Cell("A4").Value = "some text";
        ClassicAssert.AreEqual(false, ws.Evaluate("OR(FALSE,A4)"));
    }
}
