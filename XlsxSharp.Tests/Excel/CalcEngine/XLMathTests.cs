using XlsxSharp.Excel.CalcEngine.Functions;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class XlMathTests
{
    [Test]
    public void IsEven()
    {
        ClassicAssert.IsTrue(XLMath.IsEven(2));
        ClassicAssert.IsFalse(XLMath.IsEven(3));
    }

    [Test]
    public void IsOdd()
    {
        ClassicAssert.IsTrue(XLMath.IsOdd(3));
        ClassicAssert.IsFalse(XLMath.IsOdd(2));
    }
}
