using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class CalcEngineExceptionTests
{
    // Culture is reset to en-US before every test by GlobalHooks.ResetCulture, so the
    // OneTimeSetUp this class used to have here is no longer needed.

    [Test]
    public void InvalidCharNumber()
    {
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("CHAR(-2)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("CHAR(270)"));
    }

    [Test]
    public void DivisionByZero()
    {
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("0/0"));
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            new XLWorkbook().AddWorksheet().Evaluate("0/0")
        );
    }

    [Test]
    public void InvalidFunction()
    {
        ClassicAssert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("XXX(A1:A2)"));

        IXLWorksheet ws = new XLWorkbook().AddWorksheet();
        ClassicAssert.AreEqual(XLError.NameNotRecognized, ws.Evaluate("XXX(A1:A2)"));
    }

    [Test]
    public void NestedNameNotRecognizedException()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").SetFormulaA1("=XXX");
        ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

        ClassicAssert.AreEqual("Success", ws.Cell("A2").Value);
    }
}
