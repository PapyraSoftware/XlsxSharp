using System;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.Misc;

public class FormulaTests
{
    [Test]
    public void CopyFormula()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").FormulaA1 = "B1";
        ws.Cell("A1").CopyTo("A2");
        ClassicAssert.AreEqual("B2", ws.Cell("A2").FormulaA1);
    }

    [Test]
    public void CopyFormula2()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").FormulaA1 = "A2-1";
            ws.Cell("A1").CopyTo("B1");
            ClassicAssert.AreEqual("R[1]C-1", ws.Cell("A1").FormulaR1C1);
            ClassicAssert.AreEqual("R[1]C-1", ws.Cell("B1").FormulaR1C1);
            ClassicAssert.AreEqual("B2-1", ws.Cell("B1").FormulaA1);

            ws.Cell("A1").FormulaA1 = "B1+1";
            ws.Cell("A1").CopyTo("A2");
            ClassicAssert.AreEqual("RC[1]+1", ws.Cell("A1").FormulaR1C1);
            ClassicAssert.AreEqual("RC[1]+1", ws.Cell("A2").FormulaR1C1);
            ClassicAssert.AreEqual("B2+1", ws.Cell("A2").FormulaA1);
        }
    }

    [Test]
    public void CopyFormulaWithSheetNameThatResemblesFormula()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("S10 Data");
            ws.Cell("A1").Value = "Some value";
            ws.Cell("A2").Value = 123;

            ws = wb.Worksheets.Add("Summary");
            ws.Cell("A1").FormulaA1 = "='S10 Data'!A1";
            ClassicAssert.AreEqual("Some value", ws.Cell("A1").Value);

            ws.Cell("A1").CopyTo("A2");
            ClassicAssert.AreEqual("'S10 Data'!A2", ws.Cell("A2").FormulaA1);

            ws.Cell("A1").CopyTo("B1");
            ClassicAssert.AreEqual("'S10 Data'!B1", ws.Cell("B1").FormulaA1);

            ws.Cell("A3").FormulaA1 = "=SUM('S10 Data'!A2)";
            ClassicAssert.AreEqual(123, ws.Cell("A3").Value);
        }
    }

    [Test]
    public void FormulaWithReferenceIncludingSheetName()
    {
        using (XLWorkbook wb = new())
        {
            object value;
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
            ws.Cell("B1").FormulaA1 = "=SUM(A1:A50)";
            value = ws.Cell("B1").Value;
            ClassicAssert.AreEqual(1275, value);

            ws = wb.AddWorksheet("Sheet2");

            ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet1!A50)";
            value = ws.Cell("A1").Value;
            ClassicAssert.AreEqual(1275, value);

            ws.Cell("B1").FormulaA1 = "=SUM(Sheet1!A1:A50)";
            value = ws.Cell("B1").Value;
            ClassicAssert.AreEqual(1275, value);
        }
    }

    [Test]
    public void InvalidReferences()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
            ws = wb.AddWorksheet("Sheet2");

            ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet2!A50)";
            ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("A1").Value);

            ws.Cell("B1").FormulaA1 = "=SUM(UnknownSheet!A50)";
            ClassicAssert.AreEqual(XLError.CellReference, ws.Cell("B1").Value);
        }
    }

    [Test]
    public void DateAgainstStringComparison()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = new DateTime(2016, 1, 1);

            ws.Cell("A2").FormulaA1 = @"=IF(A1 = """", ""A"", ""B"")";
            XLCellValue actual = ws.Cell("A2").Value;
            ClassicAssert.AreEqual(actual, "B");

            ws.Cell("A3").FormulaA1 = @"=IF("""" = A1, ""A"", ""B"")";
            actual = ws.Cell("A3").Value;
            ClassicAssert.AreEqual(actual, "B");
        }
    }

    [Test]
    public void FormulaThatReferencesEntireRow()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().Value = 1;
            ws.FirstCell().CellRight().Value = 2;
            ws.FirstCell().CellRight(5).Value = 3;

            ws.FirstCell().CellBelow().FormulaA1 = "=SUM(1:1)";

            XLCellValue actual = ws.FirstCell().CellBelow().Value;
            ClassicAssert.AreEqual(6, actual);
        }
    }

    [Test]
    public void FormulaThatReferencesEntireColumn()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().Value = 1;
            ws.FirstCell().CellBelow().Value = 2;
            ws.FirstCell().CellBelow(5).Value = 3;

            ws.FirstCell().CellRight().FormulaA1 = "=SUM(A:A)";

            XLCellValue actual = ws.FirstCell().CellRight().Value;
            ClassicAssert.AreEqual(6, actual);
        }
    }

    [Test]
    public void FormulaThatStartsWithEqualsAndPlus()
    {
        object actual;
        actual = XLWorkbook.EvaluateExpr("=MID(\"This is a test\", 6, 2)");
        ClassicAssert.AreEqual("is", actual);

        actual = XLWorkbook.EvaluateExpr("=+MID(\"This is a test\", 6, 2)");
        ClassicAssert.AreEqual("is", actual);

        actual = XLWorkbook.EvaluateExpr("=+++++MID(\"This is a test\", 6, 2)");
        ClassicAssert.AreEqual("is", actual);

        actual = XLWorkbook.EvaluateExpr("+MID(\"This is a test\", 6, 2)");
        ClassicAssert.AreEqual("is", actual);
    }

    [Test]
    public void UnimplementedStandardFunctionsAreEvaluatedToNameNotFoundError()
    {
        // RTD will never be implemented
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            "RTD(\"MyRTDServerProdID\",\"MyServer\",\"RaceNum\",\"RunnerID\",\"StatType\")"
        );
        ClassicAssert.AreEqual(XLError.NameNotRecognized, actual);
    }

    [Test]
    public void FormulasWithErrors()
    {
        ClassicAssert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("YEAR(#REF!)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("YEAR(#VALUE!)"));
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("YEAR(#DIV/0!)"));
        ClassicAssert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("YEAR(#NAME?)"));
        ClassicAssert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr("YEAR(#N/A)"));
        ClassicAssert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("YEAR(#NULL!)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("YEAR(#NUM!)"));
    }

    [Test]
    public void LegacyFunctionPropagateErrorWithoutException() =>
        ClassicAssert.AreEqual(
            XLError.NameNotRecognized,
            XLWorkbook.EvaluateExpr("SIN(YEAR(#NAME?))+1")
        );

    [Test]
    public void UnicodeLetterParsing()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet("Sheet C CÄ");
            IXLWorksheet ws2 = wb.AddWorksheet("ÖC");
            IXLWorksheet ws3 = wb.AddWorksheet("Sheet3");

            ws1.FirstCell().SetValue(100);
            ws2.FirstCell().SetValue(50);

            ws3.FirstCell().FormulaA1 = "='Sheet C CÄ'!A1";
            ws3.FirstCell().CellBelow().FormulaA1 = "ÖC!A1";

            ClassicAssert.AreEqual(100, ws3.FirstCell().Value);
            ClassicAssert.AreEqual(50, ws3.FirstCell().CellBelow().Value);
        }
    }

    [Test, Skip("Shifting formulas is done by regexp that breaks array formula.")]
    public void ShiftFormula()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.Cell("B1").FormulaA1 = "ATAN2(C1,C2)";
            ws.Cell("B2").FormulaA1 = "DEC2HEX(C2)";
            ws.Range("B3:B5").FormulaArrayA1 = "DAYS360(C3:C5, D3:D5)";

            ws.Column(1).Delete();

            ClassicAssert.AreEqual("ATAN2(B1,B2)", ws.Cell("A1").FormulaA1);
            ClassicAssert.AreEqual("DEC2HEX(B2)", ws.Cell("A2").FormulaA1);
            ClassicAssert.True(ws.Cell("A3").HasArrayFormula);
            ClassicAssert.AreEqual("DAYS360(B3:B5, C3:C5)", ws.Cell("A3").FormulaA1);
        }
    }
}
