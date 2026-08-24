using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class FunctionsTests
{
    [Test]
    public void Asc()
    {
        object actual;

        actual = XLWorkbook.EvaluateExpr(@"Asc(""Text"")");
        ClassicAssert.AreEqual("Text", actual);
    }

    [Test]
    public void Clean()
    {
        object actual;

        actual = XLWorkbook.EvaluateExpr(string.Format(@"Clean(""A{0}B"")", Environment.NewLine));
        ClassicAssert.AreEqual("AB", actual);
    }

    [Test]
    public void Dollar()
    {
        using XLWorkbook wb = new();
        object actual = wb.Evaluate("DOLLAR(12345.123)");
        ClassicAssert.AreEqual(TestHelper.CurrencySymbol + "12,345.12", actual);

        actual = wb.Evaluate("DOLLAR(12345.123, 1)");
        ClassicAssert.AreEqual(TestHelper.CurrencySymbol + "12,345.1", actual);
    }

    [Test]
    [Arguments("A", "A", true)]
    [Arguments("A", "a", false)]
    [Arguments("", "", true)]
    public void Exact(string lhs, string rhs, bool result)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"EXACT(\"{lhs}\", \"{rhs}\")");
        ClassicAssert.AreEqual(result, actual);
    }

    [Test]
    public void ExactConvertsValuesToText()
    {
        ClassicAssert.AreEqual(false, XLWorkbook.EvaluateExpr("EXACT(TRUE, \"true\")"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("EXACT(TRUE, \"TRUE\")"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("EXACT(1, \"1\")"));
        ClassicAssert.AreEqual(true, XLWorkbook.EvaluateExpr("EXACT(IF(TRUE,), \"\")"));

        // Check blank cell
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(true, ws.Evaluate("EXACT(A1, \"\")"));
    }

    [Test]
    public void ExactPropagatesErrors()
    {
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExpr("EXACT(#DIV/0!, \"A\")")
        );
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExpr("EXACT(\"A\", #DIV/0!)")
        );
    }

    [Test]
    public void Fixed()
    {
        object actual;

        actual = XLWorkbook.EvaluateExpr("Fixed(12345.123)");
        ClassicAssert.AreEqual("12,345.12", actual);

        actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1)");
        ClassicAssert.AreEqual("12,345.1", actual);

        actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1, TRUE)");
        ClassicAssert.AreEqual("12345.1", actual);
    }

    [Test]
    public void FormulaFromAnotherSheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet("ws1");
        ws1.FirstCell().SetValue(1).CellRight().SetFormulaA1("A1 + 1");
        IXLWorksheet ws2 = wb.AddWorksheet("ws2");
        ws2.FirstCell().SetFormulaA1("ws1!B1 + 1");
        object v = ws2.FirstCell().Value;
        ClassicAssert.AreEqual(3.0, v);
    }

    [Test]
    public void TextConcat()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 1;
        ws.Cell("B1").Value = 1;
        ws.Cell("B2").Value = 1;

        ws.Cell("C1").FormulaA1 = "\"The total value is: \" & SUM(A1:B2)";

        object r = ws.Cell("C1").Value;
        ClassicAssert.AreEqual("The total value is: 4", r);
    }

    [Test]
    public void Trim()
    {
        ClassicAssert.AreEqual("Test", XLWorkbook.EvaluateExpr("Trim(\"Test    \")"));

        //Should not trim non breaking space
        //See http://office.microsoft.com/en-us/excel-help/trim-function-HP010062581.aspx
        ClassicAssert.AreEqual("Test\u00A0", XLWorkbook.EvaluateExpr("Trim(\"Test\u00A0 \")"));
    }

    [Test]
    public void TestEmptyTallyOperations()
    {
        //In these test no values have been set
        XLWorkbook wb = new();
        wb.Worksheets.Add("TallyTests");
        IXLCell cell = wb.Worksheet(1).Cell(1, 1).SetFormulaA1("=MAX(D1,D2)");
        ClassicAssert.AreEqual(0, cell.Value);
        cell = wb.Worksheet(1).Cell(2, 1).SetFormulaA1("=MIN(D1,D2)");
        ClassicAssert.AreEqual(0, cell.Value);
        cell = wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=SUM(D1,D2)");
        ClassicAssert.AreEqual(0, cell.Value);
    }

    [Test]
    public void TestOmittedParameters()
    {
        using (XLWorkbook wb = new())
        {
            object value;
            value = wb.Evaluate("=IF(TRUE,1)");
            ClassicAssert.AreEqual(1, value);

            value = wb.Evaluate("=IF(TRUE,1,)");
            ClassicAssert.AreEqual(1, value);

            value = wb.Evaluate("=ISBLANK(IF(FALSE,1,))");
            ClassicAssert.AreEqual(true, value);

            value = wb.Evaluate("=IF(FALSE,,2)");
            ClassicAssert.AreEqual(2, value);
        }
    }

    [Test]
    public void TestDefaultExcelFunctionNamespace()
    {
        ClassicAssert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("TODAY()"));
        ClassicAssert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("_xlfn.TODAY()"));
        ClassicAssert.IsTrue((bool)XLWorkbook.EvaluateExpr("_xlfn.TODAY() = TODAY()"));
    }

    [Test]
    [Arguments("=1234%", 12.34)]
    [Arguments("=1234%%", 0.1234)]
    [Arguments("=100+200%", 102.0)]
    [Arguments("=100%+200", 201.0)]
    [Arguments("=(100+200)%", 3.0)]
    [Arguments("=200%^5", 32.0)]
    [Arguments("=200%^400%", 16.0)]
    [Arguments("=SUM(100,200,300)%", 6.0)]
    public void PercentOperator(string formula, double expectedResult)
    {
        double res = (double)XLWorkbook.EvaluateExpr(formula);

        ClassicAssert.AreEqual(expectedResult, res, XLHelper.Epsilon);
    }

    [Test]
    [Arguments("=--1", 1)]
    [Arguments("=++1", 1)]
    [Arguments("=-+-+-1", -1)]
    [Arguments("=2^---2", 0.25)]
    public void MultipleUnaryOperators(string formula, double expectedResult)
    {
        double res = (double)XLWorkbook.EvaluateExpr(formula);

        ClassicAssert.AreEqual(expectedResult, res, XLHelper.Epsilon);
    }

    [Test]
    [Arguments("RIGHT(\"2020\", 2) + 1", 21)]
    [Arguments("LEFT(\"20.2020\", 6) + 1", 21.202)]
    [Arguments("2 + (\"3\" & \"4\")", 36)]
    [Arguments("2 + \"3\" & \"4\"", "54")]
    [Arguments("\"7\" & \"4\"", "74")]
    public void TestStringSubExpression(string formula, object expectedResult)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(formula);

        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    public void CellFunctionIsEvaluatedToReferenceError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "$B$4(5)";

        ClassicAssert.AreEqual(XLError.CellReference, ws.Cell("A1").Value);
    }

    [Test]
    [Arguments("BASE(1E15,30)", "MathTrig.Base")]
    [Arguments("COMBIN(1000,500)", "MathTrig.Combin")]
    [Arguments("COMBINA(100,500)", "MathTrig.CombinA")]
    [Arguments("DECIMAL(\"ZZZ\",26)", "MathTrig.Decimal")]
    [Arguments("GCD(123456,54124)", "MathTrig.Gcd")]
    [Arguments("LCM(123456,54124)", "MathTrig.Lcm")]
    [Arguments("MDETERM({1,2;3,4})", "MathTrig.MDeterm")]
    [Arguments("MINVERSE({1,2;3,4})", "MathTrig.MInverse")]
    [Arguments("MMULT({1},{2})", "MathTrig.MMult")]
    [Arguments("MULTINOMIAL(2,3,4)", "MathTrig.Multinomial")]
    [Arguments("PRODUCT(2,3,{4,5,6})", "MathTrig.Product")]
    [Arguments("SERIESSUM(2,1,2,{1,2,3})", "MathTrig.SeriesSum")]
    [Arguments("SUM(2,1,2,{1,2,3})", "MathTrig.Sum")]
    [Arguments("SUMIF(B1:B4,\"=5\")", "MathTrig.SumIf")]
    [Arguments("SUMIFS(B1:B4,C1:C4,\">0\")", "MathTrig.SumIfs")]
    [Arguments("SUMPRODUCT({2,3},{4,5})", "MathTrig.SumProduct")]
    [Arguments("SUMSQ(5,4)", "MathTrig.SumSq")]
    [Arguments("NETWORKDAYS(10,100,{20,50})", "DateAndTime.NetWorkDays")]
    [Arguments("WORKDAY(10,100,{20,50})", "DateAndTime.Workday")]
    [Arguments("YEARFRAC(1,10000,1)", "DateAndTime.YearFrac")]
    [Arguments("AND(TRUE,TRUE)", "Logical.And")]
    [Arguments("OR(TRUE,TRUE)", "Logical.Or")]
    [Arguments("COLUMN(D:Z)", "Lookup.Column")]
    [Arguments("HLOOKUP(2,{0,1,2,3},1)", "Lookup.Hlookup")]
    [Arguments("MATCH(5,{1,7})", "Lookup.Match")]
    [Arguments("ROW(2:5)", "Lookup.Row")]
    [Arguments("VLOOKUP(2,{0;1;2;3},1)", "Lookup.Vlookup")]
    [Arguments("AVERAGE({1,2;3,4})", "Statistical.Average")]
    [Arguments("AVERAGEA({1,2;3,4})", "Statistical.AverageA")]
    [Arguments("BINOMDIST(6,10,0.5,TRUE)", "Statistical.BinomDist")]
    [Arguments("BINOM.DIST(6,10,0.5,TRUE)", "Statistical.BinomDist")]
    [Arguments("COUNT({1,2})", "Statistical.Count")]
    [Arguments("COUNTA({1,2})", "Statistical.Count")]
    [Arguments("COUNTBLANK(D1:D7)", "Statistical.CountBlank")]
    [Arguments("COUNTIF(D1:D7,\">0\")", "Statistical.CountIf")]
    [Arguments("COUNTIFS(D1:D7,\">0\")", "Statistical.CountIfs")]
    [Arguments("DEVSQ({1,2})", "Statistical.DevSq")]
    [Arguments("GEOMEAN({1,2})", "Statistical.GeoMean")]
    [Arguments("LARGE({10,11,12},2)", "Statistical.Large")]
    [Arguments("MAX({1,5})", "Statistical.Max")]
    [Arguments("MAXA({1,5})", "Statistical.MaxA")]
    [Arguments("MEDIAN({1,5,7})", "Statistical.Median")]
    [Arguments("MIN({1,5})", "Statistical.Min")]
    [Arguments("MINA({1,5})", "Statistical.MinA")]
    [Arguments("STDEV({1,5})", "Statistical.StDev")]
    [Arguments("STDEVA({1,5})", "Statistical.StDevA")]
    [Arguments("STDEVP({1,5})", "Statistical.StDevP")]
    [Arguments("STDEVPA({1,5})", "Statistical.StDevPA")]
    [Arguments("STDEV.S({1,5})", "Statistical.StDev")]
    [Arguments("STDEV.P({1,5})", "Statistical.StDevP")]
    [Arguments("VAR({1,5})", "Statistical.Var")]
    [Arguments("VARA({1,5})", "Statistical.VarA")]
    [Arguments("VARP({1,5})", "Statistical.VarP")]
    [Arguments("VARPA({1,5})", "Statistical.VarPA")]
    [Arguments("VAR.S({1,5})", "Statistical.Var")]
    [Arguments("VAR.P({1,5})", "Statistical.VarP")]
    [Arguments("ASC(\"A\")", "Text.Asc")]
    [Arguments("CLEAN(\"A\")", "Text.Clean")]
    [Arguments("CONCAT(\"A\",\"B\")", "Text.Concat")]
    [Arguments("CONCATENATE(\"A\",\"B\")", "Text.Concatenate")]
    [Arguments("LEFT(\"AB\",1)", "Text.Left")]
    [Arguments("LOWER(\"A\")", "Text.Lower")]
    [Arguments("NUMBERVALUE(\"1\")", "Text.NumberValue")]
    [Arguments("PROPER(\"hello\")", "Text.Proper")]
    [Arguments("REPT(\"A\",10)", "Text.Rept")]
    [Arguments("RIGHT(\"AB\",1)", "Text.Right")]
    [Arguments("T({100})", "Text.T")]
    [Arguments("TEXTJOIN(\"-\",TRUE,\"A\",\"B\")", "Text.TextJoin")]
    [Arguments("TRIM(\"ABC\")", "Text.Trim")]
    public void CanCancelFunctionExecution(string formula, string expectedStackTrace)
    {
        CancellationTokenSource cts = new();
        using XLWorkbook wb = new(new LoadOptions { CancellationToken = cts.Token });
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("D1").Value = 4; // Need a non-blank cell for COUNTBLANK check
        ws.Cell("A1").FormulaA1 = formula;

        cts.Cancel();
        OperationCanceledException? ex = ClassicAssert.Throws<OperationCanceledException>(() =>
            _ = ws.Cell("A1").Value
        );

        StringAssert.Contains(expectedStackTrace + "(", ex?.StackTrace);
    }
}
