using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class CompareOperatorsTests
{
    [Test]
    [Arguments("1=1", true)]
    [Arguments("1=0", false)]
    [Arguments("0.0=0", true)]
    [Arguments("TRUE=TRUE", true)]
    [Arguments("FALSE=FALSE", true)]
    [Arguments("TRUE=FALSE", false)]
    [Arguments("\"text\"=\"text\"", true)]
    [Arguments("\"tExT\"=\"TeXt\"", true)]
    [Arguments("\"text\"=\"text\"", true)]
    [Arguments("\"\"=\"\"", true)]
    [Arguments("#VALUE!=#VALUE!", XLError.IncompatibleValue)]
    [Arguments("A1=B1", true)] // blanks are equal
    public void EqualToWithSameType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    [Test]
    [Arguments("1<>1", false)]
    [Arguments("1<>0", true)]
    [Arguments("0.0<>0", false)]
    [Arguments("TRUE<>TRUE", false)]
    [Arguments("FALSE<>FALSE", false)]
    [Arguments("TRUE<>FALSE", true)]
    [Arguments("\"texty\"<>\"text\"", true)]
    [Arguments("\"tExT\"<>\"TeXt\"", false)]
    [Arguments("\"text\"<>\"text\"", false)]
    [Arguments("\"\"<>\"\"", false)]
    [Arguments("#VALUE!<>#VALUE!", XLError.IncompatibleValue)]
    [Arguments("A1<>B1", false)] // blanks are equal
    public void NotEqualToWithSameType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    [Test]
    [Arguments("1>1", false)]
    [Arguments("1>0", true)]
    [Arguments("0.0>0", false)]
    [Arguments("TRUE>TRUE", false)]
    [Arguments("FALSE>FALSE", false)]
    [Arguments("TRUE>FALSE", true)]
    [Arguments("\"text\">\"text\"", false)]
    [Arguments("\"texu\">\"text\"", true)]
    [Arguments("#VALUE!>#REF!", XLError.IncompatibleValue)]
    [Arguments("A1>A2", false)]
    public void GreaterThenWithSameType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    [Test]
    [Arguments("1>=1", true)]
    [Arguments("1>=0", true)]
    [Arguments("0.0>=0", true)]
    [Arguments("TRUE>=TRUE", true)]
    [Arguments("FALSE>=FALSE", true)]
    [Arguments("TRUE>=FALSE", true)]
    [Arguments("\"text\">=\"text\"", true)]
    [Arguments("\"texu\">=\"text\"", true)]
    [Arguments("#VALUE!>=#REF!", XLError.IncompatibleValue)]
    [Arguments("A1>=A2", true)]
    public void GreaterThenOrEqualWithSameType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    [Test]
    [Arguments("-5<5", true)]
    [Arguments("1<1", false)]
    [Arguments("1<0", false)]
    [Arguments("0.0<0", false)]
    [Arguments("TRUE<TRUE", false)]
    [Arguments("FALSE<FALSE", false)]
    [Arguments("TRUE<FALSE", false)]
    [Arguments("FALSE<TRUE", true)]
    [Arguments("\"text\"<\"text\"", false)]
    [Arguments("\"text\"<\"texu\"", true)]
    [Arguments("#VALUE!<#REF!", XLError.IncompatibleValue)]
    [Arguments("A1<A2", false)]
    public void LessThenWithSameType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    [Test]
    [Arguments("-5<=5", true)]
    [Arguments("1<=1", true)]
    [Arguments("1<=0", false)]
    [Arguments("0.0<=0", true)]
    [Arguments("TRUE<=TRUE", true)]
    [Arguments("FALSE<=FALSE", true)]
    [Arguments("TRUE<=FALSE", false)]
    [Arguments("FALSE<=TRUE", true)]
    [Arguments("\"text\"<=\"text\"", true)]
    [Arguments("\"text\"<=\"texu\"", true)]
    [Arguments("#VALUE!<=#REF!", XLError.IncompatibleValue)]
    [Arguments("A1<=A2", true)]
    public void LessThenOrEqualWithSameType(string formula, object expectedValue) =>
        ClassicAssert.AreEqual(expectedValue, Evaluate(formula));

    [Test]
    [Arguments("TRUE>-1", true)]
    [Arguments("TRUE>1", true)]
    [Arguments("TRUE>100", true)]
    [Arguments("FALSE>-1", true)]
    [Arguments("FALSE>1", true)]
    [Arguments("FALSE>100", true)]
    [Arguments("TRUE>\"100\"", true)]
    [Arguments("FALSE>\"100\"", true)]
    [Arguments("FALSE>\"\"", true)]
    [Arguments("\"\">FALSE", false)]
    [Arguments("10>FALSE", false)]
    [Arguments("10>TRUE", false)]
    [Arguments("-1<TRUE", true)]
    [Arguments("1<TRUE", true)]
    [Arguments("100<TRUE", true)]
    [Arguments("-1<FALSE", true)]
    [Arguments("1<FALSE", true)]
    [Arguments("100<FALSE", true)]
    [Arguments("\"100\"<TRUE", true)]
    [Arguments("\"100\"<FALSE", true)]
    [Arguments("\"\"<FALSE", true)]
    [Arguments("FALSE<\"\"", false)]
    [Arguments("FALSE<10", false)]
    [Arguments("TRUE<10", false)]
    public void ComparisonLogicalIsAlwaysGreaterThanAnyTextOrNumber(
        string formula,
        bool expectedResult
    ) => ClassicAssert.AreEqual(expectedResult, Evaluate(formula));

    [Test]
    [Arguments("\"\">10", true)]
    [Arguments("\"1\">10", true)]
    [Arguments("10<\"\"", true)]
    [Arguments("10<\"1\"", true)]
    public void ComparisonTextIsAlwaysGreaterThanAnyNumber(string formula, bool expectedResult) =>
        ClassicAssert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("FALSE=A1")]
    [Arguments("A1=FALSE")]
    [Arguments("A1=0")]
    [Arguments("0=A1")]
    [Arguments("\"\"=A1")]
    [Arguments("A1=\"\"")]
    public void ComparisonBlankIsEqualToFalseOrZeroOrEmptyString(string formula) =>
        ClassicAssert.AreEqual(true, Evaluate(formula));

    private static XLCellValue Evaluate(string formula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        return ws.Evaluate(formula);
    }
}
