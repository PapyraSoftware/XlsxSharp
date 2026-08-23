using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class MathTrigTests
{
    private const double Tolerance = 1e-10;

    // Replicates NUnit's [Range(from, to, step)]/[Range(from, to)] parameter data generators,
    // which this class used heavily before the move to TUnit (which has no direct equivalent).
    internal static IEnumerable<double> DoubleRange(double from, double to, double step)
    {
        int count = (int)((to - from) / step) + 1;
        for (int i = 0; i < count; i++)
        {
            yield return from + (step * i);
        }
    }

    internal static IEnumerable<double> Range0To10Step01() => DoubleRange(0d, 10d, 0.1);

    internal static IEnumerable<double> RangeNeg10ToNeg01Step01() => DoubleRange(-10d, -0.1, 0.1);

    internal static IEnumerable<double> Range11To3Step01() => DoubleRange(1.1, 3d, 0.1);

    internal static IEnumerable<double> RangeNeg1To09Step01() => DoubleRange(-1d, 0.9, 0.1);

    internal static IEnumerable<double> RangeNeg09To09Step01() => DoubleRange(-0.9, 0.9, 0.1);

    internal static IEnumerable<double> RangeNeg3ToNeg11Step01() => DoubleRange(-3d, -1.1, 0.1);

    internal static IEnumerable<double> Range01To5Step04() => DoubleRange(0.1, 5d, 0.4);

    internal static IEnumerable<double> RangeNeg5ToNeg01Step03() => DoubleRange(-5d, -0.1, 0.3);

    internal static IEnumerable<double> RangeNeg5ToNeg01Step04() => DoubleRange(-5d, -0.1, 0.4);

    internal static IEnumerable<double> Range01To5Step03() => DoubleRange(0.1, 5d, 0.3);

    internal static IEnumerable<double> Range1To5Step02() => DoubleRange(1d, 5d, 0.2);

    internal static IEnumerable<int> IntRange(int from, int to)
    {
        for (int value = from; value <= to; value++)
        {
            yield return value;
        }
    }

    [Test]
    [MethodDataSource(nameof(Range0To10Step01))]
    public void AbsReturnsItselfOnPositiveNumbers(double input)
    {
        double actual = (double)
            XLWorkbook.EvaluateExpr(
                string.Format(@"ABS({0})", input.ToString(CultureInfo.InvariantCulture))
            );
        ClassicAssert.AreEqual(input, actual, Tolerance * 10);
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg10ToNeg01Step01))]
    public void AbsReturnsTheCorrectValueOnNegativeInput(double input)
    {
        double actual = (double)
            XLWorkbook.EvaluateExpr(
                string.Format(@"ABS({0})", input.ToString(CultureInfo.InvariantCulture))
            );
        ClassicAssert.AreEqual(-input, actual, Tolerance * 10);
    }

    [Test]
    [Arguments(-1, 3.141592654)]
    [Arguments(-0.9, 2.690565842)]
    [Arguments(-0.8, 2.498091545)]
    [Arguments(-0.7, 2.346193823)]
    [Arguments(-0.6, 2.214297436)]
    [Arguments(-0.5, 2.094395102)]
    [Arguments(-0.4, 1.982313173)]
    [Arguments(-0.3, 1.875488981)]
    [Arguments(-0.2, 1.772154248)]
    [Arguments(-0.1, 1.670963748)]
    [Arguments(0, 1.570796327)]
    [Arguments(0.1, 1.470628906)]
    [Arguments(0.2, 1.369438406)]
    [Arguments(0.3, 1.266103673)]
    [Arguments(0.4, 1.159279481)]
    [Arguments(0.5, 1.047197551)]
    [Arguments(0.6, 0.927295218)]
    [Arguments(0.7, 0.79539883)]
    [Arguments(0.8, 0.643501109)]
    [Arguments(0.9, 0.451026812)]
    [Arguments(1, 0)]
    public void AcosReturnsCorrectValue(double input, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ACOS({input})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance * 10);
    }

    [Test]
    [MethodDataSource(nameof(Range11To3Step01))]
    public void AcosReturnsErrorWhenNumberOutsideRange(double input)
    {
        // checking input and it's additive inverse as both are outside range.
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOS({input})"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOS({-input})"));
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg1To09Step01))]
    public void AcoshNumbersBelow1ThrowNumberException(double input) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOSH({input})"));

    [Test]
    [Arguments(1.2, 0.622362504)]
    [Arguments(1.5, 0.96242365)]
    [Arguments(1.8, 1.192910731)]
    [Arguments(2.1, 1.372859144)]
    [Arguments(2.4, 1.522079367)]
    [Arguments(2.7, 1.650193455)]
    [Arguments(3, 1.762747174)]
    [Arguments(3.3, 1.863279351)]
    [Arguments(3.6, 1.954207529)]
    [Arguments(3.9, 2.037266466)]
    [Arguments(4.2, 2.113748231)]
    [Arguments(4.5, 2.184643792)]
    [Arguments(4.8, 2.250731414)]
    [Arguments(5.1, 2.312634419)]
    [Arguments(5.4, 2.370860342)]
    [Arguments(5.7, 2.425828318)]
    [Arguments(6, 2.47788873)]
    [Arguments(1, 0)]
    public void AcoshReturnsCorrectNumber(double angle, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ACOSH({angle})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance * 10);
    }

    [Test]
    [Arguments(-10, 3.041924001)]
    [Arguments(-9, 3.030935432)]
    [Arguments(-8, 3.017237659)]
    [Arguments(-7, 2.999695599)]
    [Arguments(-6, 2.976443976)]
    [Arguments(-5, 2.944197094)]
    [Arguments(-4, 2.89661399)]
    [Arguments(-3, 2.819842099)]
    [Arguments(-2, 2.677945045)]
    [Arguments(-1, 2.35619449)]
    [Arguments(0, 1.570796327)]
    [Arguments(1, 0.785398163)]
    [Arguments(2, 0.463647609)]
    [Arguments(3, 0.321750554)]
    [Arguments(4, 0.244978663)]
    [Arguments(5, 0.19739556)]
    [Arguments(6, 0.165148677)]
    [Arguments(7, 0.141897055)]
    [Arguments(8, 0.124354995)]
    [Arguments(9, 0.110657221)]
    [Arguments(10, 0.099668652)]
    public void AcotReturnsCorrectNumber(double angle, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ACOT({angle})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance * 10);
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg09To09Step01))]
    public void AcothReturnsErrorForAbsoluteAngleSmallerThanOne(double angle) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOTH({angle})"));

    [Test]
    [Arguments(-10, -0.100335348)]
    [Arguments(-9, -0.111571776)]
    [Arguments(-8, -0.125657214)]
    [Arguments(-7, -0.143841036)]
    [Arguments(-6, -0.168236118)]
    [Arguments(-5, -0.202732554)]
    [Arguments(-4, -0.255412812)]
    [Arguments(-3, -0.34657359)]
    [Arguments(-2, -0.549306144)]
    [Arguments(2, 0.549306144)]
    [Arguments(3, 0.34657359)]
    [Arguments(4, 0.255412812)]
    [Arguments(5, 0.202732554)]
    [Arguments(6, 0.168236118)]
    [Arguments(7, 0.143841036)]
    [Arguments(8, 0.125657214)]
    [Arguments(9, 0.111571776)]
    [Arguments(10, 0.100335348)]
    [Arguments(1E+100, 1E-100)]
    public void AcothReturnsCorrectNumber(double angle, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ACOTH({angle})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance * 10);
    }

    [Test]
    [Arguments("LVII", 57)]
    [Arguments(@"mcmxii", 1912)]
    [Arguments("", 0)]
    [Arguments("-IV", -4)]
    [Arguments("   XIV   ", 14)]
    [Arguments(@"MCMLXXXIII ", 1983)]
    [Arguments(@"IIIIIIIIM", 992)]
    [Arguments(@"CIVIIX", 102)]
    [Arguments(@"IIX", 8)]
    [Arguments(@"VIII", 8)]
    public void ArabicReturnsCorrectNumber(string roman, int arabic)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ARABIC(\"{roman}\")");
        ClassicAssert.AreEqual(arabic, actual);
    }

    [Test]
    public void ArabicSolitaryMinusIsNotValidRomanNumber() =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("ARABIC(\"-\")"));

    [Test]
    public void ArabicCanHaveAtMost255Chars()
    {
        ClassicAssert.AreEqual(
            255000,
            XLWorkbook.EvaluateExpr($"ARABIC(\"{new string('M', 255)}\")")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"ARABIC(\"{new string('M', 256)}\")")
        );
    }

    [Test]
    [Arguments("- I")]
    [Arguments("roman")]
    public void ArabicReturnsConversionErrorOnInvalidNumbers(string invalidRoman) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"ARABIC(\"{invalidRoman}\")")
        );

    [Test]
    [Arguments(-1, -1.570796327)]
    [Arguments(-0.9, -1.119769515)]
    [Arguments(-0.8, -0.927295218)]
    [Arguments(-0.7, -0.775397497)]
    [Arguments(-0.6, -0.643501109)]
    [Arguments(-0.5, -0.523598776)]
    [Arguments(-0.4, -0.411516846)]
    [Arguments(-0.3, -0.304692654)]
    [Arguments(-0.2, -0.201357921)]
    [Arguments(-0.1, -0.100167421)]
    [Arguments(0, 0)]
    [Arguments(0.1, 0.100167421)]
    [Arguments(0.2, 0.201357921)]
    [Arguments(0.3, 0.304692654)]
    [Arguments(0.4, 0.411516846)]
    [Arguments(0.5, 0.523598776)]
    [Arguments(0.6, 0.643501109)]
    [Arguments(0.7, 0.775397497)]
    [Arguments(0.8, 0.927295218)]
    [Arguments(0.9, 1.119769515)]
    [Arguments(1, 1.570796327)]
    public void AsinReturnsCorrectResult(double input, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ASIN({input})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance * 10);
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg3ToNeg11Step01))]
    public void AsinThrowsNumberExceptionWhenAbsOfInputGreaterThan1(double input)
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ASIN({input})"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ASIN({-input})"));
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.1, 0.0998340788992076)]
    [Arguments(0.2, 0.198690110349241)]
    [Arguments(0.3, 0.295673047563422)]
    [Arguments(0.4, 0.390035319770715)]
    [Arguments(0.5, 0.481211825059603)]
    [Arguments(0.6, 0.568824898732248)]
    [Arguments(0.7, 0.652666566082356)]
    [Arguments(0.8, 0.732668256045411)]
    [Arguments(0.9, 0.808866935652783)]
    [Arguments(1, 0.881373587019543)]
    [Arguments(2, 1.44363547517881)]
    [Arguments(3, 1.81844645923207)]
    [Arguments(4, 2.0947125472611)]
    [Arguments(5, 2.31243834127275)]
    public void AsinhReturnsCorrectResult(double input, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ASINH({input})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
        double minusActual = (double)XLWorkbook.EvaluateExpr($"ASINH({-input})");
        ClassicAssert.AreEqual(-expectedResult, minusActual, Tolerance);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.1, 0.099668652491162)]
    [Arguments(0.2, 0.197395559849881)]
    [Arguments(0.3, 0.291456794477867)]
    [Arguments(0.4, 0.380506377112365)]
    [Arguments(0.5, 0.463647609000806)]
    [Arguments(0.6, 0.540419500270584)]
    [Arguments(0.7, 0.610725964389209)]
    [Arguments(0.8, 0.674740942223553)]
    [Arguments(0.9, 0.732815101786507)]
    [Arguments(1, 0.785398163397448)]
    [Arguments(2, 1.10714871779409)]
    [Arguments(3, 1.24904577239825)]
    [Arguments(4, 1.32581766366803)]
    [Arguments(5, 1.37340076694502)]
    public void AtanReturnsCorrectResult(double input, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN({input})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
        double minusActual = (double)XLWorkbook.EvaluateExpr($"ATAN({-input})");
        ClassicAssert.AreEqual(-expectedResult, minusActual, Tolerance);
    }

    [Test]
    [MethodDataSource(nameof(Range01To5Step04))]
    public void Atan2Returns0OnSecond0AndFirstGreater0(double input)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, 0)");
        ClassicAssert.AreEqual(0, actual, Tolerance);
    }

    [Test]
    [Arguments(1, 2, 1.10714871779409)]
    [Arguments(1, 3, 1.24904577239825)]
    [Arguments(2, 3, 0.98279372324733)]
    [Arguments(1, 4, 1.32581766366803)]
    [Arguments(3, 4, 0.92729521800161)]
    [Arguments(1, 5, 1.37340076694502)]
    [Arguments(2, 5, 1.19028994968253)]
    [Arguments(3, 5, 1.03037682652431)]
    [Arguments(4, 5, 0.89605538457134)]
    [Arguments(1, 6, 1.40564764938027)]
    [Arguments(5, 6, 0.87605805059819)]
    [Arguments(1, 7, 1.42889927219073)]
    [Arguments(2, 7, 1.29249666778979)]
    [Arguments(3, 7, 1.16590454050981)]
    [Arguments(4, 7, 1.05165021254837)]
    [Arguments(5, 7, 0.95054684081208)]
    [Arguments(6, 7, 0.86217005466723)]
    public void Atan2ReturnsCorrectResultsEqualOnAllMultiplesOfFraction(
        double x,
        double y,
        double expectedResult
    )
    {
        for (int i = 1; i < 5; i++)
        {
            double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({x * i}, {y * i})");
            ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
        }
    }

    [Test]
    [MethodDataSource(nameof(Range01To5Step04))]
    public void Atan2ReturnsHalfPiOn0AsFirstInputWhenSecondGreater0(double input)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2(0, {input})");
        ClassicAssert.AreEqual(0.5 * Math.PI, actual, Tolerance);
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg5ToNeg01Step03))]
    public void Atan2ReturnsMinus3QuartersOfPiWhenFirstSmaller0AndSecondItsNegative(double input)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, {input})");
        ClassicAssert.AreEqual(-0.75 * Math.PI, actual, Tolerance);
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg5ToNeg01Step04))]
    public void Atan2ReturnsMinusHalfPiOn0AsFirstInputWhenSecondSmaller0(double input)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2(0, {input})");
        ClassicAssert.AreEqual(-0.5 * Math.PI, actual, Tolerance);
    }

    [Test]
    [MethodDataSource(nameof(RangeNeg5ToNeg01Step04))]
    public void Atan2ReturnsPiOn0AsSecondInputWhenFirstSmaller0(double input)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, 0)");
        ClassicAssert.AreEqual(Math.PI, actual, Tolerance);
    }

    [Test]
    [MethodDataSource(nameof(Range01To5Step03))]
    public void Atan2ReturnsQuarterOfPiWhenInputsAreEqualAndGreater0(double input)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, {input})");
        ClassicAssert.AreEqual(0.25 * Math.PI, actual, Tolerance);
    }

    [Test]
    public void Atan2ThrowsDiv0ExceptionOn0And0() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("ATAN2(0, 0)"));

    [Test]
    [Arguments(-0.99, -2.64665241236225)]
    [Arguments(-0.9, -1.47221948958322)]
    [Arguments(-0.8, -1.09861228866811)]
    [Arguments(-0.6, -0.693147180559945)]
    [Arguments(-0.4, -0.423648930193602)]
    [Arguments(-0.2, -0.202732554054082)]
    [Arguments(0, 0)]
    [Arguments(0.2, 0.202732554054082)]
    [Arguments(0.4, 0.423648930193602)]
    [Arguments(0.6, 0.693147180559945)]
    [Arguments(0.8, 1.09861228866811)]
    [Arguments(-0.9, -1.47221948958322)]
    [Arguments(-0.990, -2.64665241236225)]
    [Arguments(-0.999, -3.8002011672502)]
    public void AtanhReturnsCorrectResults(double input, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"ATANH({input})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance * 10);
    }

    [Test]
    [MethodDataSource(nameof(Range1To5Step02))]
    public void AtanhThrowsNumberExceptionWhenAbsOfInput1OrGreater(double input)
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ATANH({input})"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ATANH({-input})"));
    }

    [Test]
    [Arguments(0, 36, "0")]
    [Arguments(1, 36, "1")]
    [Arguments(2, 36, "2")]
    [Arguments(3, 36, "3")]
    [Arguments(4, 36, "4")]
    [Arguments(5, 36, "5")]
    [Arguments(6, 36, "6")]
    [Arguments(7, 36, "7")]
    [Arguments(8, 36, "8")]
    [Arguments(9, 36, "9")]
    [Arguments(10, 36, "A")]
    [Arguments(11, 36, "B")]
    [Arguments(12, 36, "C")]
    [Arguments(13, 36, "D")]
    [Arguments(14, 36, "E")]
    [Arguments(15, 36, "F")]
    [Arguments(16, 36, "G")]
    [Arguments(17, 36, "H")]
    [Arguments(18, 36, "I")]
    [Arguments(19, 36, "J")]
    [Arguments(20, 36, "K")]
    [Arguments(21, 36, "L")]
    [Arguments(22, 36, "M")]
    [Arguments(23, 36, "N")]
    [Arguments(24, 36, "O")]
    [Arguments(25, 36, "P")]
    [Arguments(26, 36, "Q")]
    [Arguments(27, 36, "R")]
    [Arguments(28, 36, "S")]
    [Arguments(29, 36, "T")]
    [Arguments(30, 36, "U")]
    [Arguments(31, 36, "V")]
    [Arguments(32, 36, "W")]
    [Arguments(33, 36, "X")]
    [Arguments(34, 36, "Y")]
    [Arguments(35, 36, "Z")]
    [Arguments(36, 36, "10")]
    [Arguments(255, 29, "8N")]
    [Arguments(255, 2, "11111111")]
    public void BaseReturnsNumberInSpecifiedBase(int input, int radix, string expectedResult)
    {
        string actual = (string)XLWorkbook.EvaluateExpr($"BASE({input},{radix})");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Arguments(255, 2, 3, "11111111")]
    [Arguments(255, 2, 8, "11111111")]
    [Arguments(255, 2, 10, "0011111111")]
    [Arguments(10, 3, 4, "0101")]
    [Arguments(0, 10, 0, "")]
    public void BaseReturnsTextOfAtLeastMinimalLength(
        int input,
        int radix,
        int minLength,
        string expectedResult
    )
    {
        string actual = (string)XLWorkbook.EvaluateExpr($"BASE({input},{radix},{minLength})");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    public void BaseMinLengthMustBeAtMost255() =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("BASE(0,2,256)"));

    [Test]
    [Arguments(@"""x""", "2", "2")]
    [Arguments("0", @"""x""", "2")]
    [Arguments("0", "2", @"""x""")]
    public void BaseCoercion(string input, string radix, string minLength) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"BASE({input},{radix},{minLength})")
        );

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [-2, 1])]
    [MethodDataSource(nameof(IntRange), Arguments = [37, 40])]
    public void BaseRadixMustBeBetween2And36(int radix) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"BASE(0,{radix})"));

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [-5, -1])]
    public void BaseNumberMustBeZeroOrPositive(int input) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"BASE({input},2)"));

    [Test]
    public void BaseNumberMustFitInDoubleWithoutPrecisionLoss()
    {
        ClassicAssert.AreEqual(@"2GOPQOE5GCG", XLWorkbook.EvaluateExpr("BASE(9.007E+15,36)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("BASE(9.008E+15,36)")
        );
    }

    [Test]
    [Arguments(24.3, 5, 25)]
    [Arguments(6.7, 1, 7)]
    [Arguments(-8.1, 2, -8)]
    [Arguments(5.5, 2.1, 6.3)]
    [Arguments(5.5, 0, 0)]
    [Arguments(-5.5, 2.1, -4.2)]
    [Arguments(-5.5, -2.1, -6.3)]
    [Arguments(-5.5, 0, 0)]
    [Arguments(0, 0, 0)]
    [Arguments(0, 0.1, 0)]
    [Arguments(0, -0.1, 0)]
    [Arguments(0.1, 0, 0)]
    [Arguments(-0.1, 0, 0)]
    public void Ceiling(double input, double significance, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"CEILING({input}, {significance})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
    }

    [Test]
    [Arguments(6.7, -1)]
    [Arguments(0.1, -0.2)]
    public void CeilingReturnsErrorOnDifferentNumberAndSignificance(
        double input,
        double significance
    ) =>
        // Spec says "if x and significance have different signs, #NUM! is returned.",
        // but in reality it only happens when number is positive and step negative.
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"CEILING({input}, {significance})")
        );

    [Test]
    [Arguments(24.3, 5, null, 25)]
    [Arguments(6.7, null, null, 7)]
    [Arguments(-8.1, 2, null, -8)]
    [Arguments(-5.5, 2, -1, -6)]
    [Arguments(-5.5, 2, -0.1, -6)]
    [Arguments(5.5, 2.1, 0, 6.3)]
    [Arguments(5.5, -2.1, 0, 6.3)]
    [Arguments(5.5, 0, 0, 0)]
    [Arguments(5.5, 2.1, -1, 6.3)]
    [Arguments(5.5, -2.1, -1, 6.3)]
    [Arguments(5.5, 0, -1, 0)]
    [Arguments(5.5, 2.1, 10, 6.3)]
    [Arguments(5.5, -2.1, 10, 6.3)]
    [Arguments(5.5, 0, 10, 0)]
    [Arguments(-5.5, 2.1, 0, -4.2)]
    [Arguments(-5.5, -2.1, 0, -4.2)]
    [Arguments(-5.5, 0, 0, 0)]
    [Arguments(-5.5, 2.1, -1, -6.3)]
    [Arguments(-5.5, -2.1, -1, -6.3)]
    [Arguments(-5.5, 0, -1, 0)]
    [Arguments(-5.5, 2.1, 10, -6.3)]
    [Arguments(-5.5, -2.1, 10, -6.3)]
    [Arguments(-5.5, 0, 10, 0)]
    public void CeilingMath(double input, double? significance, double? mode, double expectedResult)
    {
        StringBuilder parameters = new();
        parameters.Append(input);
        if (significance != null)
        {
            parameters.Append(", ").Append(significance);
            if (mode != null)
            {
                parameters.Append(", ").Append(mode);
            }
        }

        double actual = (double)XLWorkbook.EvaluateExpr($"CEILING.MATH({parameters})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
    }

    [Test]
    public void Combin()
    {
        XLCellValue actual1 = XLWorkbook.EvaluateExpr("COMBIN(200, 2)");
        ClassicAssert.AreEqual(19900.0, actual1);

        XLCellValue actual2 = XLWorkbook.EvaluateExpr("COMBIN(20.1, 2.9)");
        ClassicAssert.AreEqual(190.0, actual2);
    }

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [0, 10])]
    public void CombinReturns1ForKIs0OrKEqualsN(int n)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, 0)");
        ClassicAssert.AreEqual(1, actual);

        XLCellValue actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n})");
        ClassicAssert.AreEqual(1, actual2);
    }

    [Test]
    [Arguments(0, 0, 1)]
    [Arguments(1, 0, 1)]
    [Arguments(1, 1, 1)]
    [Arguments(4, 2, 6)]
    [Arguments(5, 2, 10)]
    [Arguments(6, 2, 15)]
    [Arguments(6, 3, 20)]
    [Arguments(7, 2, 21)]
    [Arguments(7, 3, 35)]
    public void CombinCalculatesCombinations(int n, int k, int expectedResult)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, {k})");
        ClassicAssert.AreEqual(expectedResult, actual);

        XLCellValue actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n - k})");
        ClassicAssert.AreEqual(expectedResult, actual2);
    }

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [1, 10])]
    public void CombinReturnsNForKIs1OrKIsNMinus1(int n)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, 1)");
        ClassicAssert.AreEqual(n, actual);

        XLCellValue actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n - 1})");
        ClassicAssert.AreEqual(n, actual2);
    }

    [Test]
    public void CombinReturnsNumErrorWhenKIsLargerThanN()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(5, 6)"));

        // Values are floored, so this is COMBIN(5, 5).
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("COMBIN(5, 5.5)"));
    }

    [Test]
    public void CombinReturnsNumErrorWhenValueIsTooLarge()
    {
        // Maximum int - 1 is maximum computable value in Excel.
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("COMBIN(2147483647, 2147483647)")
        );
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(5E+301, 6)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(6, 5E+301)"));
    }

    [Test]
    [Arguments(-4)]
    [Arguments(-3)]
    [Arguments(-1)]
    [Arguments(-0.1)]
    public void CombinReturnsNumErrorForAnyArgumentSmallerThan0(double smaller0)
    {
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, {1})",
                    smaller0.ToString(CultureInfo.InvariantCulture),
                    (-smaller0).ToString(CultureInfo.InvariantCulture)
                )
            )
        );

        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, {1})",
                    (-smaller0).ToString(CultureInfo.InvariantCulture),
                    smaller0.ToString(CultureInfo.InvariantCulture)
                )
            )
        );
    }

    [Test]
    [Arguments("\"no number\"")]
    [Arguments("\"\"")]
    public void CombinReturnsValueErrorForAnyNonNumericArgument(string input)
    {
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"COMBIN({input}, 1)")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"COMBIN(1, {input})")
        );
    }

    [Test]
    [Arguments(4, 3, 20)]
    [Arguments(10, 3, 220)]
    [Arguments(0, 0, 1)]
    [Arguments(1, 0, 1)]
    [Arguments(10, 15, 1307504)]
    public void CombinaCalculatesCorrectValues(int number, int chosen, int expectedResult)
    {
        XLCellValue actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
        ClassicAssert.AreEqual(expectedResult, actualResult);
    }

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [0, 10])]
    public void CombinaReturnsOneWhenChosenIsZero(int number)
    {
        XLCellValue actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, 0)");
        ClassicAssert.AreEqual(1, actualResult);
    }

    [Test]
    [Arguments(-1, 2)]
    [Arguments(-3, -2)]
    [Arguments(2, -2)]
    [Arguments(int.MaxValue + 1d, 1)]
    public void CombinaReturnsErrorOnInvalidValues(double number, int chosen) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})")
        );

    [Test]
    [Arguments(4.23, 3, 20)]
    [Arguments(10.4, 3.14, 220)]
    [Arguments(0, 0.4, 1)]
    public void CombinaTruncatesNumbersToZero(double number, double chosen, int expectedResult)
    {
        XLCellValue actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
        ClassicAssert.AreEqual(expectedResult, actualResult);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(0.4, 0.921060994002885)]
    [Arguments(0.8, 0.696706709347165)]
    [Arguments(1.2, 0.362357754476674)]
    [Arguments(1.6, -0.0291995223012888)]
    [Arguments(2, -0.416146836547142)]
    [Arguments(2.4, -0.737393715541245)]
    [Arguments(2.8, -0.942222340668658)]
    [Arguments(3.2, -0.998294775794753)]
    [Arguments(3.6, -0.896758416334147)]
    [Arguments(4, -0.653643620863612)]
    [Arguments(4.4, -0.307332869978419)]
    [Arguments(4.8, 0.0874989834394464)]
    [Arguments(5.2, 0.468516671300377)]
    [Arguments(5.6, 0.77556587851025)]
    [Arguments(6, 0.960170286650366)]
    [Arguments(6.4, 0.993184918758193)]
    [Arguments(6.8, 0.869397490349825)]
    [Arguments(7.2, 0.608351314532255)]
    [Arguments(7.6, 0.251259842582256)]
    [Arguments(8, -0.145500033808614)]
    [Arguments(8.4, -0.519288654116686)]
    public void CosReturnsCorrectResult(double input, double expectedResult)
    {
        double actualResult = (double)XLWorkbook.EvaluateExpr($"COS({input})");
        ClassicAssert.AreEqual(expectedResult, actualResult, Tolerance);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(0.4, 1.08107237183845)]
    [Arguments(0.8, 1.33743494630484)]
    [Arguments(1.2, 1.81065556732437)]
    [Arguments(1.6, 2.57746447119489)]
    [Arguments(2, 3.76219569108363)]
    [Arguments(2.4, 5.55694716696551)]
    [Arguments(2.8, 8.25272841686113)]
    [Arguments(3.2, 12.2866462005439)]
    [Arguments(3.6, 18.3127790830626)]
    [Arguments(4, 27.3082328360165)]
    [Arguments(4.4, 40.7315730024356)]
    [Arguments(4.8, 60.7593236328919)]
    [Arguments(5.2, 90.638879219786)]
    [Arguments(5.6, 135.215052644935)]
    [Arguments(6, 201.715636122456)]
    [Arguments(6.4, 300.923349714678)]
    [Arguments(6.8, 448.924202712783)]
    [Arguments(7.2, 669.715755490113)]
    [Arguments(7.6, 999.098197777775)]
    [Arguments(8, 1490.47916125218)]
    [Arguments(8.4, 2223.53348628359)]
    public void CoshReturnsCorrectResult(double input, double expectedResult)
    {
        double actualResult = (double)XLWorkbook.EvaluateExpr($"COSH({input})");
        ClassicAssert.AreEqual(expectedResult, actualResult, Tolerance);
        double actualResult2 = (double)XLWorkbook.EvaluateExpr($"COSH({-input})");
        ClassicAssert.AreEqual(expectedResult, actualResult2, Tolerance);
    }

    [Test]
    [Arguments(711)]
    [Arguments(-711)]
    [Arguments(100000)]
    public void CoshTooLargeReturnsError(double input) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"COSH({input})"));

    [Test]
    [Arguments(1, 0.642092616)]
    [Arguments(2, -0.457657554)]
    [Arguments(3, -7.015252551)]
    [Arguments(4, 0.863691154)]
    [Arguments(5, -0.295812916)]
    [Arguments(6, -3.436353004)]
    [Arguments(7, 1.147515422)]
    [Arguments(8, -0.147065064)]
    [Arguments(9, -2.210845411)]
    [Arguments(10, 1.542351045)]
    [Arguments(11, -0.004425741)]
    [Arguments(Math.PI * 0.5, 0)]
    [Arguments(45, 0.617369624)]
    [Arguments(-2, 0.457657554)]
    [Arguments(-3, 7.015252551)]
    public void Cot(double angle, double expected)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"COT({angle})");
        ClassicAssert.AreEqual(expected, actual, Tolerance * 10.0);
    }

    [Test]
    public void CotReturnsDivisionByZeroErrorOnAngleZero() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("COT(0)"));

    [Test]
    public void CothReturnsDivisionByZeroErrorOnAngleZero() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("COTH(0)"));

    [Test]
    [Arguments(-10, -1.000000004)]
    [Arguments(-9, -1.00000003)]
    [Arguments(-8, -1.000000225)]
    [Arguments(-7, -1.000001663)]
    [Arguments(-6, -1.000012289)]
    [Arguments(-5, -1.000090804)]
    [Arguments(-4, -1.00067115)]
    [Arguments(-3, -1.004969823)]
    [Arguments(-2, -1.037314721)]
    [Arguments(-1, -1.313035285)]
    [Arguments(1, 1.313035285)]
    [Arguments(2, 1.037314721)]
    [Arguments(3, 1.004969823)]
    [Arguments(4, 1.00067115)]
    [Arguments(5, 1.000090804)]
    [Arguments(6, 1.000012289)]
    [Arguments(7, 1.000001663)]
    [Arguments(8, 1.000000225)]
    [Arguments(9, 1.00000003)]
    [Arguments(10, 1.000000004)]
    public void CothReturnsCorrectNumber(double angle, double expected)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"COTH({angle})");
        ClassicAssert.AreEqual(expected, actual, Tolerance * 10.0);
    }

    [Test]
    public void CscReturnsDivisionByZeroOnAngleZero() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("CSC(0)"));

    [Test]
    [Arguments(-10, 1.838163961)]
    [Arguments(-9, -2.426486644)]
    [Arguments(-8, -1.010756218)]
    [Arguments(-7, -1.522101063)]
    [Arguments(-6, 3.578899547)]
    [Arguments(-5, 1.042835213)]
    [Arguments(-4, 1.321348709)]
    [Arguments(-3, -7.086167396)]
    [Arguments(-2, -1.09975017)]
    [Arguments(-1, -1.188395106)]
    [Arguments(1, 1.188395106)]
    [Arguments(2, 1.09975017)]
    [Arguments(3, 7.086167396)]
    [Arguments(4, -1.321348709)]
    [Arguments(5, -1.042835213)]
    [Arguments(6, -3.578899547)]
    [Arguments(7, 1.522101063)]
    [Arguments(8, 1.010756218)]
    [Arguments(9, 2.426486644)]
    [Arguments(10, -1.838163961)]
    public void CscReturnsCorrectNumber(double angle, double expected)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"CSC({angle})");
        ClassicAssert.AreEqual(expected, actual, Tolerance * 10);
    }

    [Test]
    [Arguments(1, 0.850918128)]
    [Arguments(2, 0.275720565)]
    [Arguments(3, 0.09982157)]
    [Arguments(4, 0.03664357)]
    [Arguments(5, 0.013476506)]
    [Arguments(6, 0.004957535)]
    [Arguments(7, 0.001823765)]
    [Arguments(8, 0.000670925)]
    [Arguments(9, 0.00024682)]
    [Arguments(10, 0.000090799859712122200000)]
    [Arguments(11, 0.0000334034)]
    public void CschCalculatesCorrectValues(double input, double expectedOutput) =>
        ClassicAssert.AreEqual(
            expectedOutput,
            (double)XLWorkbook.EvaluateExpr($"CSCH({input})"),
            0.000000001
        );

    [Test]
    public void CschReturnsDivisionErrorOnAngleZero() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("CSCH(0)"));

    [Test]
    [Arguments("FF", 16, 255)]
    [Arguments("111", 2, 7)]
    [Arguments("zap", 36, 45745)] // Case insensitive
    [Arguments("  1234", 10, 1234)] // Trims start
    [Arguments("123", 10.9, 123)] // Radix truncated
    [Arguments("1F", 10, XLError.NumberInvalid)]
    [Arguments("", 10, 0)]
    public void Decimal(string inputString, double radix, object expectedResult)
    {
        XLCellValue actualResult = XLWorkbook.EvaluateExpr($"DECIMAL(\"{inputString}\", {radix})");
        ClassicAssert.AreEqual(expectedResult, actualResult);
    }

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [37, 255])]
    [MethodDataSource(nameof(IntRange), Arguments = [-5, 1])]
    public void DecimalRadixMustBeBetween2And36(int radix) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})")
        );

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [2, 36])]
    public void DecimalZeroIsZeroInAnyRadix(int radix) =>
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));

    [Test]
    public void DecimalTextMustBeLessThan256CharsLong()
    {
        string text = new('0', 256);
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"DECIMAL(\"{text}\", 10)")
        );
    }

    [Test]
    public void DecimalReturnsNumberInvalidWhenResultOutOfBounds()
    {
        ClassicAssert.AreEqual(
            1.4057081148316923E+308d,
            (double)XLWorkbook.EvaluateExpr($"DECIMAL(\"{new string('Z', 198)}\", 36)")
        );
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"DECIMAL(\"{new string('Z', 199)}\", 36)")
        );
    }

    [Test]
    [Arguments("101", "\"1 2/2\"", 5)] // 101 in binary is 5
    public void DecimalCoercion(string input, string radix, object expectedResult) =>
        ClassicAssert.AreEqual(
            expectedResult,
            XLWorkbook.EvaluateExpr($"DECIMAL({input}, {radix})")
        );

    [Test]
    public void Degrees()
    {
        double actual = (double)XLWorkbook.EvaluateExpr("DEGREES(PI())");
        ClassicAssert.AreEqual(180, actual, XLHelper.Epsilon);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(Math.PI, 180)]
    [Arguments(Math.PI * 2, 360)]
    [Arguments(1, 57.2957795130823)]
    [Arguments(2, 114.591559026165)]
    [Arguments(3, 171.887338539247)]
    [Arguments(4, 229.183118052329)]
    [Arguments(5, 286.478897565412)]
    [Arguments(6, 343.774677078494)]
    [Arguments(7, 401.070456591576)]
    [Arguments(8, 458.366236104659)]
    [Arguments(9, 515.662015617741)]
    [Arguments(10, 572.957795130823)]
    [Arguments(Math.PI * 0.5, 90)]
    [Arguments(Math.PI * 1.5, 270)]
    [Arguments(Math.PI * 0.25, 45)]
    [Arguments(-1, -57.2957795130823)]
    public void DegreesReturnsCorrectResult(double input, double expected)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"DEGREES({input})");
        ClassicAssert.AreEqual(expected, actual, Tolerance);
    }

    [Test]
    [Arguments(3, 4)]
    [Arguments(2, 2)]
    [Arguments(-1, -2)]
    [Arguments(-2, -2)]
    [Arguments(0, 0)]
    [Arguments(1.5, 2)]
    [Arguments(2.01, 4)]
    [Arguments(1e+100, 1e+100)]
    [Arguments(Math.PI, 4)]
    public void Even(double number, double expectedResult)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"EVEN({number})");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(1, Math.E)]
    [Arguments(2, 7.38905609893065)]
    [Arguments(3, 20.0855369231877)]
    [Arguments(4, 54.5981500331442)]
    [Arguments(5, 148.413159102577)]
    [Arguments(6, 403.428793492735)]
    [Arguments(7, 1096.63315842846)]
    [Arguments(8, 2980.95798704173)]
    [Arguments(9, 8103.08392757538)]
    [Arguments(10, 22026.4657948067)]
    [Arguments(11, 59874.1417151978)]
    [Arguments(12, 162754.791419004)]
    [Arguments(-1E+100, 0)]
    public void ExpReturnsCorrectResults(double input, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"EXP({input})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
    }

    [Test]
    [Arguments(710)]
    public void ExpWithTooLargeResultReturnError(double input) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"EXP({input})"));

    [Test]
    public void Fact()
    {
        object actual = XLWorkbook.EvaluateExpr("Fact(5.9)");
        ClassicAssert.AreEqual(120.0, actual);
    }

    [Test]
    [Arguments(0, 1d)]
    [Arguments(1, 1d)]
    [Arguments(2, 2d)]
    [Arguments(3, 6d)]
    [Arguments(4, 24d)]
    [Arguments(5, 120d)]
    [Arguments(6, 720d)]
    [Arguments(7, 5040d)]
    [Arguments(8, 40320d)]
    [Arguments(9, 362880d)]
    [Arguments(10, 3628800d)]
    [Arguments(11, 39916800d)]
    [Arguments(12, 479001600d)]
    [Arguments(13, 6227020800d)]
    [Arguments(14, 87178291200d)]
    [Arguments(15, 1307674368000d)]
    [Arguments(16, 20922789888000d)]
    [Arguments(170.9, 7.257415615308004E+306)]
    [Arguments(0.1, 1L)]
    [Arguments(2.3, 2L)]
    [Arguments(2.8, 2L)]
    public void FactCalculatesFactorial(double input, double expectedResult)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            $@"FACT({input.ToString(CultureInfo.InvariantCulture)})"
        );
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Arguments(-10)]
    [Arguments(-5)]
    [Arguments(-1)]
    [Arguments(-0.1)]
    public void FactReturnsErrorForNegativeInput(double input)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            $@"FACT({input.ToString(CultureInfo.InvariantCulture)})"
        );
        ClassicAssert.AreEqual(XLError.NumberInvalid, actual);
    }

    [Test]
    [Arguments(171)]
    [Arguments(5000)]
    public void FactReturnsErrorForTooLargeResult(int input)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($@"FACT({input})");
        ClassicAssert.AreEqual(XLError.NumberInvalid, actual);
    }

    [Test]
    public void FactCoercionFailsForNonNumericInput() =>
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"FACT(""x"")"));

    [Test]
    [Arguments(0, 1L)]
    [Arguments(1, 1L)]
    [Arguments(2, 2L)]
    [Arguments(3, 3L)]
    [Arguments(4, 8L)]
    [Arguments(5, 15L)]
    [Arguments(6, 48L)]
    [Arguments(7, 105L)]
    [Arguments(8, 384L)]
    [Arguments(9, 945L)]
    [Arguments(10, 3840L)]
    [Arguments(11, 10395L)]
    [Arguments(12, 46080L)]
    [Arguments(13, 135135L)]
    [Arguments(14, 645120)]
    [Arguments(15, 2027025)]
    [Arguments(16, 10321920)]
    [Arguments(-1, 1L)]
    [Arguments(0, 1)]
    [Arguments(0.1, 1L)]
    [Arguments(1.4, 1L)]
    [Arguments(2.3, 2L)]
    [Arguments(2.8, 2L)]
    public void FactDoubleReturnsCorrectResult(double input, long expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"FACTDOUBLE({input})");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Arguments(301)]
    [Arguments(1e+100)]
    public void FactDoubleReturnsErrorOnTooLargeValue(double n) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"FACTDOUBLE({n})"));

    [Test]
    [MethodDataSource(nameof(IntRange), Arguments = [-10, -2])]
    public void FactDoubleThrowsNumberExceptionForInputSmallerThanMinus1(int input) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"FACTDOUBLE({input})")
        );

    [Test]
    public void FactDoubleThrowsValueExceptionForNonNumericInput() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FACTDOUBLE(""x"")")
        );

    [Test]
    [Arguments(0, 0, 0)]
    [Arguments(0, 1, 0)]
    [Arguments(24.3, 5, 20)]
    [Arguments(6.7, 1, 6)]
    [Arguments(-8.1, 2, -10)]
    [Arguments(5.5, 2.1, 4.2)]
    [Arguments(-5.5, 2.1, -6.3)]
    [Arguments(-5.5, -2.1, -4.2)]
    public void Floor(double input, double significance, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
    }

    [Test]
    [Arguments(6.7, 0)]
    [Arguments(-6.7, 0)]
    public void FloorThrowsDivisionByZeroOnZeroSignificance(double input, double significance) =>
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})")
        );

    [Test]
    [Arguments(6.7, -1)]
    public void FloorThrowsNumberExceptionOnInvalidInput(double input, double significance) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})")
        );

    [Test]
    // Functions have to support a period first before we can implement this
    [Arguments(24.3, 5, null, 20)]
    [Arguments(6.7, null, null, 6)]
    [Arguments(-8.1, 2, null, -10)]
    [Arguments(5.5, 2.1, 0, 4.2)]
    [Arguments(5.5, -2.1, 0, 4.2)]
    [Arguments(5.5, 0, 0, 0)]
    [Arguments(5.5, 2.1, -1, 4.2)]
    [Arguments(5.5, -2.1, -1, 4.2)]
    [Arguments(5.5, 0, -2, 0)]
    [Arguments(5.5, 2.1, 10, 4.2)]
    [Arguments(5.5, -2.1, 10, 4.2)]
    [Arguments(5.5, 0, 10, 0)]
    [Arguments(-5.5, 2.1, 0, -6.3)]
    [Arguments(-5.5, -2.1, 0, -6.3)]
    [Arguments(-5.5, 0, 0, 0)]
    [Arguments(-5.5, 2.1, -1, -4.2)]
    [Arguments(-5.5, -2.1, -1, -4.2)]
    [Arguments(-5.5, 0, -1, 0)]
    [Arguments(-5.5, 2.1, 10, -4.2)]
    [Arguments(-5.5, -2.1, 10, -4.2)]
    [Arguments(-5.5, 0, 0, 0)]
    public void FloorMath(double input, double? significance, int? mode, double expectedResult)
    {
        StringBuilder parameters = new();
        parameters.Append(input);
        if (significance != null)
        {
            parameters.Append(", ").Append(significance);
            if (mode != null)
            {
                parameters.Append(", ").Append(mode);
            }
        }

        double actual = (double)XLWorkbook.EvaluateExpr($"FLOOR.MATH({parameters})");
        ClassicAssert.AreEqual(expectedResult, actual, Tolerance);
    }

    [Test]
    [Arguments("24,36", 12)]
    [Arguments("240,360,30", 30)]
    [Arguments("24.9,36.9", 12)]
    [Arguments("{24,36}", 12)]
    [Arguments("{\"24\",\"36\"}", 12)]
    [Arguments("5,0", 5)]
    [Arguments("0,5", 5)]
    public void Gcd(string args, double expected)
    {
        ClassicAssert.AreEqual(expected, (double)XLWorkbook.EvaluateExpr($"GCD({args})"));
    }

    [Test]
    public void GcdAcceptsReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[] { (120, 240), ("60", "150") });
        ClassicAssert.AreEqual(30, ws.Evaluate("GCD(A1:A2,B1:B2)"));

        // Blank is considered 0
        ClassicAssert.AreEqual(60, ws.Evaluate("GCD(A1:A3)"));

        // Logical are not converted
        ws.Cell("A3").Value = true;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("GCD(A1:A3)"));

        // Unconvertable text causes error
        ws.Cell("A3").Value = "one";
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("GCD(A1:A3)"));
    }

    [Test]
    [Arguments]
    public void GcdNumbersMustFitInDoubleWithoutPrecisionLoss()
    {
        ClassicAssert.AreEqual(9.007E+15, XLWorkbook.EvaluateExpr("GCD(9.007E+15)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("GCD(9.008E+15)"));
    }

    [Test]
    [Arguments]
    public void GcdNumbersMustBeZeroOrPositive() =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("GCD(-1)"));

    [Test]
    [Arguments(8.9, 8)]
    [Arguments(-8.9, -9)]
    public void Int(double input, double expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"INT({input})");
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Arguments("24, 36", 72)]
    [Arguments("24.9, 36.9", 72)]
    [Arguments("{24, 36}", 72)]
    [Arguments("{1,2,3;4,5,6}", 60)]
    [Arguments("{\"1\",\"2\",\"3\"}", 6)]
    [Arguments("240, 360, 30", 720)]
    [Arguments("5, 0", 0)]
    [Arguments("0, 5", 0)]
    public void Lcm(string args, double expected)
    {
        ClassicAssert.AreEqual(expected, (double)XLWorkbook.EvaluateExpr($"LCM({args})"));
    }

    [Test]
    public void LcmAcceptsReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[] { (1, 2, 3), ("4", "5", "6") });
        ClassicAssert.AreEqual(60, ws.Evaluate("LCM(A1:B2,C1:C2)"));

        // Blank is considered 0
        ClassicAssert.AreEqual(0, ws.Evaluate("LCM(A1:A3)"));

        // Logical are not converted
        ws.Cell("A3").Value = true;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("LCM(A1:A3)"));

        // Unconvertable text causes error
        ws.Cell("A3").Value = "one";
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("LCM(A1:A3)"));
    }

    [Test]
    [Arguments]
    public void LcmNumbersMustFitInDoubleWithoutPrecisionLoss()
    {
        ClassicAssert.AreEqual(9.007E+15, XLWorkbook.EvaluateExpr("LCM(9.007E+15)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LCM(9.008E+15)"));
    }

    [Test]
    [Arguments]
    public void LcmNumbersMustBeZeroOrPositive() =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LCM(-1)"));

    [Test]
    [Arguments(86, 4.4543472962)]
    [Arguments(2.7182818, 0.9999999895)]
    [Arguments(20.085536923, 3)]
    public void LnCalculatesLogarithm(double x, double ln) =>
        ClassicAssert.AreEqual(ln, (double)XLWorkbook.EvaluateExpr($"LN({x})"), Tolerance);

    [Test]
    [Arguments(0)]
    [Arguments(-0.7)]
    [Arguments(-10)]
    public void LnNonPositiveReturnsError(double x) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"LN({x})"));

    [Test]
    [Arguments(10, 10, 1)]
    [Arguments(8, 2, 3)]
    [Arguments(86, 2.7182818, 4.4543473428883)]
    public void LogCalculatesLogarithm(double x, double @base, double result) =>
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"LOG({x}, {@base})"),
            Tolerance
        );

    [Test]
    public void LogDefaultBaseIs10() =>
        ClassicAssert.AreEqual(2, XLWorkbook.EvaluateExpr("LOG(100)"));

    [Test]
    public void LogErrorConditions()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LOG(0)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LOG(1,0)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LOG(0,0)"));
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("LOG(10,1)"));
    }

    [Test]
    [Arguments(86, 1.93449845124)]
    [Arguments(10, 1)]
    [Arguments(1E5, 5)]
    public void Log10CalculatesLogarithm(double x, double expectedResult) =>
        ClassicAssert.AreEqual(
            expectedResult,
            (double)XLWorkbook.EvaluateExpr($"LOG10({x})"),
            Tolerance
        );

    [Test]
    [Arguments(0)]
    [Arguments(-5)]
    [Arguments(-0.5)]
    public void Log10ErrorConditions(double x) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"LOG10({x})"));

    [Test]
    public void Log10IsDetectedInsideExpression() =>
        // Because LOG10 is extracted from CellFunction, make sure it is properly read even in the middle of expression.
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("0 + LOG10(10)"));

    [Test]
    public void MDeterm()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[] { (2, 4), (3, 5) });

        ws.Cell("A5").FormulaA1 = "MDETERM(A1:B2)";
        XLCellValue actual = ws.Cell("A5").Value;
        ClassicAssert.AreEqual(-2, (double)actual, Tolerance);

        ws.Cell("A6").FormulaA1 = "SUM(A5)";
        actual = ws.Cell("A6").Value;
        ClassicAssert.AreEqual(-2, (double)actual, Tolerance);

        ws.Cell("A7").FormulaA1 = "SUM(MDETERM(A1:B2))";
        actual = ws.Cell("A7").Value;
        ClassicAssert.AreEqual(-2, (double)actual, Tolerance);
    }

    [Test]
    public void MDetermExamples()
    {
        // Examples from spec
        ClassicAssert.AreEqual(
            1,
            (double)XLWorkbook.EvaluateExpr("MDETERM({3,6,1;1,1,0;3,10,2})"),
            Tolerance
        );
        ClassicAssert.AreEqual(
            -3,
            (double)XLWorkbook.EvaluateExpr("MDETERM({3,6;1,1})"),
            Tolerance
        );

        // Example from office website
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1")
            .InsertData(
                new object[]
                {
                    ("Data", "Data", "Data", "Data"),
                    (1, 3, 8, 5),
                    (1, 3, 6, 1),
                    (1, 1, 1, 0),
                    (7, 3, 10, 2),
                }
            );
        ClassicAssert.AreEqual(88, (double)ws.Evaluate("MDETERM(A2:D5)"), Tolerance);
    }

    [Test]
    public void MDetermRequiresEqualNumberOfRowsAndColumns() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("MDETERM({1,2})")
        );

    [Test]
    public void MDetermSingularMatrixReturnsZero() =>
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExpr("MDETERM({1,2;1,2})"));

    [Test]
    public void MDetermRequiresAllArrayElementsAreNumbers()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[] { (2, 4), (3, 5) });

        ws.Cell("B2").Value = Blank.Value;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MDETERM(A1:B2)"));

        ws.Cell("B2").Value = "1";
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MDETERM(A1:B2)"));

        ws.Cell("B2").Value = true;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MDETERM(A1:B2)"));

        ws.Cell("B2").Value = XLError.NameNotRecognized;
        ClassicAssert.AreEqual(XLError.NameNotRecognized, ws.Evaluate("MDETERM(A1:B2)"));
    }

    [Test]
    public void MInverse()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[] { (1, 2, 1), (3, 4, -1), (0, 2, 0) });

        ws.Cell("A5").FormulaA1 = "MINVERSE(A1:C3)";
        XLCellValue actual = ws.Cell("A5").Value;
        ClassicAssert.AreEqual(0.25, (double)actual, Tolerance);

        ws.Cell("A6").FormulaA1 = "SUM(A5)";
        actual = ws.Cell("A6").Value;
        ClassicAssert.AreEqual(0.25, (double)actual, Tolerance);

        ws.Cell("A7").FormulaA1 = "SUM(MINVERSE(A1:C3))";
        actual = ws.Cell("A7").Value;
        ClassicAssert.AreEqual(0.5, (double)actual, Tolerance);
    }

    [Test]
    public void MInverseReturnsErrorOnSingularMatrix()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[] { (1, 2), (1, 2) });
        ClassicAssert.AreEqual(XLError.NumberInvalid, ws.Evaluate("MINVERSE(A1:B2)"));
    }

    [Test]
    public void MInverseRequiresSquareMatrix() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("MINVERSE({1,2,3;7,5,5})")
        );

    [Test]
    public void MInverseAllArrayElementsMustBeNumbers()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[] { (1, 2), (8, 4) });

        ws.Cell("B2").Value = Blank.Value;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINVERSE(A1:B2)"));

        ws.Cell("B2").Value = true;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINVERSE(A1:B2)"));

        ws.Cell("B2").Value = "1";
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINVERSE(A1:B2)"));

        ws.Cell("B2").Value = XLError.DivisionByZero;
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("MINVERSE(A1:B2)"));
    }

    [Test]
    public void MMult()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[] { (2, 4), (3, 5), (2, 4), (3, 5) });

        ws.Cell("A5").FormulaA1 = "MMULT(A1:B2, A3:B4)";
        XLCellValue actual = ws.Cell("A5").Value;
        ClassicAssert.AreEqual(16.0, actual);

        ws.Cell("A6").FormulaA1 = "SUM(A5)";
        actual = ws.Cell("A6").Value;
        ClassicAssert.AreEqual(16.0, actual);

        ws.Cell("A7").FormulaA1 = "SUM(MMULT(A1:B2, A3:B4))";
        actual = ws.Cell("A7").Value;
        ClassicAssert.AreEqual(102.0, actual);
    }

    [Test]
    public void MMultHandlesNonSquareMatrices()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1")
            .InsertData(
                new object[]
                {
                    // 2x3
                    (1, 3, 5),
                    (2, 4, 6),
                    // 3x4
                    (10, 13, 16, 19),
                    (11, 14, 17, 20),
                    (12, 15, 18, 21),
                }
            );

        // 2x4 output expected:
        // 103, 130, 157, 184
        // 136, 172, 208, 244
        ws.Cell("A6").FormulaA1 = "MMult(A1:C2, A3:D5)";
        XLCellValue actual = ws.Cell("A6").Value;
        ClassicAssert.AreEqual(103.0, actual);

        ws.Cell("A7").FormulaA1 = "Sum(MMult(A1:C2, A3:D5))";
        actual = ws.Cell("A7").Value;
        ClassicAssert.AreEqual(1334, actual);
    }

    [Test]
    [Arguments("A2:C2", "A3:C3")] // 1x3 and 1x3
    [Arguments("A2:C4", "A5:C5")] // 3x3 and 1x3
    [Arguments("A2:C5", "A6:D6")] // 3x4 and 1x4
    public void MMultArray1RowsMustMatchArray2Column(string array1Range, string array2Range)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cells($"{array1Range}").Value = 1.0;
        ws.Cells($"{array2Range}").Value = 1.0;

        ws.Cell("A1").FormulaA1 = $"MMULT({array1Range},{array2Range})";

        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("A1").Value);
    }

    [Test]
    [Arguments("")]
    [Arguments("Text")]
    public void MMultThrowsWhenCellsContainInvalidInput(string invalidInput)
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");

        // 2x3
        ws.Cell("A1").SetValue(1).CellRight().SetValue(3).CellRight().SetValue(invalidInput);
        ws.Cell("A2").SetValue(2).CellRight().SetValue(4).CellRight().SetValue(6);

        // 3x4
        ws.Cell("A3")
            .SetValue(10)
            .CellRight()
            .SetValue(13)
            .CellRight()
            .SetValue(16)
            .CellRight()
            .SetValue(19);
        ws.Cell("A4")
            .SetValue(11)
            .CellRight()
            .SetValue(14)
            .CellRight()
            .SetValue(17)
            .CellRight()
            .SetValue(20);
        ws.Cell("A5")
            .SetValue(12)
            .CellRight()
            .SetValue(15)
            .CellRight()
            .SetValue(18)
            .CellRight()
            .SetValue(21);

        ws.Cell("A6").FormulaA1 = "MMULT(A1:C2,A3:D4)";

        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("A6").Value);
    }

    [Test]
    [Arguments(1.5, 1, 0.5)]
    [Arguments(3, 2, 1)]
    [Arguments(-3, 2, 1)]
    [Arguments(-3, -2, -1)]
    [Arguments(-4.3, -0.5, -0.3)]
    [Arguments(6.9, -0.2, -0.1)]
    [Arguments(0.7, 0.6, 0.1)]
    [Arguments(6.2, 1.1, 0.7)]
    public void Mod(double x, double y, double result)
    {
        double actual = (double)XLWorkbook.EvaluateExpr($"MOD({x}, {y})");
        ClassicAssert.AreEqual(result, actual, Tolerance);
    }

    [Test]
    public void ModDivisorZeroReturnsError()
    {
        // Spec says that "If y is 0, the return value is unspecified", but Excel says #DIV/0!, so let's go with that.
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("MOD(1, 0)"));
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("MOD(0, 0)"));
    }

    [Test]
    [Arguments(10, 3, 9.0)]
    [Arguments(10.5, 3, 12.0)]
    [Arguments(10.4, 3, 9.0)]
    [Arguments(-10, -3, -9.0)]
    [Arguments(1.3, 0.2, 1.4)]
    [Arguments(5677.912288, 10, 5680.0)]
    [Arguments(5674.912288, 10, 5670.0)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.49999, 1, 0.0)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.49999, 1, 0.0)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.49999, 1, 0.0)]
    [Arguments(-13.4, -3, -12.0)]
    [Arguments(-13.5, -3, -15.0)]
    [Arguments(0.9, 0.2, 1.0)]
    [Arguments(0.89999, 0.2, 0.8)]
    [Arguments(15.5, 3, 15.0)]
    [Arguments(1.4, 0.5, 1.5)]
    [Arguments(3, 7, 0)]
    [Arguments(3, 0, 0)]
    [Arguments(0, 10, 0)]
    [Arguments(0, -5, 0)]
    public void MRound(double number, double multiple, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})"),
            1e-12
        );
    }

    [Test]
    [Arguments(123456.123, -10)]
    [Arguments(-123456.123, 5)]
    public void MRoundExceptions(double number, double multiple) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})")
        );

    [Test]
    public void Multinomial()
    {
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("MULTINOMIAL(2)"));
        ClassicAssert.AreEqual(10, XLWorkbook.EvaluateExpr("MULTINOMIAL(2,3)"));
        ClassicAssert.AreEqual(1260, XLWorkbook.EvaluateExpr("MULTINOMIAL(2,3,4)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("MULTINOMIAL(1E+100)")
        );
    }

    [Test]
    public void MultinomialAcceptsRanges()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B2").InsertData(new[] { 2, 0, 5 });
        ws.Cell("A5").InsertData(new[] { 3, 6 });

        ClassicAssert.AreEqual(3087564480d, ws.Evaluate("MULTINOMIAL(B:XFD, 2, A5:A6)"));
    }

    [Test]
    public void MultinomialDoesntAcceptNegativeValues() =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("MULTINOMIAL(5, -1)")
        );

    [Test]
    public void MultinomialCoercion()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = true;
        ws.Cell("A2").Value = 5;
        ws.Cell("A3").Value = "1 2/2";
        ws.Cell("A4").Value = "one";

        // True is not converted
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MULTINOMIAL(A1:A2)"));

        // Text is coerced
        ClassicAssert.AreEqual(21, ws.Evaluate("MULTINOMIAL(A2:A3)"));

        // Text is coerced, errors are propagates
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MULTINOMIAL(A2:A4)"));

        // Errors are propagates
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("MULTINOMIAL(5, #DIV/0!)"));
    }

    [Test]
    [Arguments(1.5, 3)]
    [Arguments(3, 3)]
    [Arguments(2, 3)]
    [Arguments(-1, -1)]
    [Arguments(-2, -3)]
    [Arguments(0, 1)]
    [Arguments(1E+100, 1E+100)]
    public void Odd(double number, double expected)
    {
        ClassicAssert.AreEqual(expected, (double)XLWorkbook.EvaluateExpr($"ODD({number})"), 1e-12);
    }

    [Test]
    public void Pi() => ClassicAssert.AreEqual(Math.PI, XLWorkbook.EvaluateExpr("PI()"));

    [Test]
    [Arguments(2, 3, 8)]
    [Arguments(2, 0.5, 1.414213562373)]
    [Arguments(-1.234, 5.0, -2.861381721051)]
    [Arguments(1.234, 5.1, 2.9221823578798)]
    public void Power(double x, double y, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"POWER({x}, {y})"),
            1e-12
        );
    }

    [Test]
    public void PowerErrors()
    {
        // Negative base and fractional exponent
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("POWER(-5, 0.5)"));

        // Spec says this should be #DIV/0!, but Excel says #NUM!
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("POWER(0, 0)"));

        // base is zero and exponent is negative -> #NUM!
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("POWER(0, -5)"));

        // Result is not representable (e.g. out fo range)
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("POWER(1e+100, 1e+100)")
        );
    }

    [Test]
    public void Product()
    {
        ClassicAssert.AreEqual(24d, XLWorkbook.EvaluateExpr("PRODUCT(2,3,4)"));

        // Examples from specification
        ClassicAssert.AreEqual(1d, XLWorkbook.EvaluateExpr("PRODUCT(1)"));
        ClassicAssert.AreEqual(120d, XLWorkbook.EvaluateExpr("PRODUCT(1,2,3,4,5)"));
        ClassicAssert.AreEqual(24d, XLWorkbook.EvaluateExpr("PRODUCT({1,2;3,4})"));
        ClassicAssert.AreEqual(120d, XLWorkbook.EvaluateExpr("PRODUCT({2,3},4,\"5\")"));

        // If no arguments are passed, return 0
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExpr("PRODUCT({\"hello\"})"));

        // Scalar blank is skipped
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("PRODUCT(IF(TRUE,), 1)"));

        // Scalar logical is converted to number
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExpr("PRODUCT(FALSE, 1)"));
        ClassicAssert.AreEqual(2, XLWorkbook.EvaluateExpr("PRODUCT(2, TRUE)"));

        // Scalar text is converted to number
        ClassicAssert.AreEqual(5, XLWorkbook.EvaluateExpr("PRODUCT(\"5\")"));

        // Scalar text that is not convertible return error
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("PRODUCT(1, \"Hello\")")
        );

        // Array non-number arguments are ignored
        ClassicAssert.AreEqual(5, XLWorkbook.EvaluateExpr("PRODUCT({5, \"Hello\", FALSE, TRUE})"));

        // Reference argument only uses number, ignores blanks, logical and text
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = 2;
        ws.Cell("A6").Value = 3;
        ClassicAssert.AreEqual(6, ws.Evaluate("PRODUCT(A1:A6)"));

        // Scalar error is propagated
        ClassicAssert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("PRODUCT(1, #NULL!)"));

        // Array error is propagated
        ClassicAssert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("PRODUCT({1, #NULL!})"));

        // Reference error is propagated
        ws.Cell("A1").Value = XLError.NoValueAvailable;
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("PRODUCT(A1)"));
    }

    [Test]
    [Arguments(5, 2, 2)]
    [Arguments(4.5, 3.1, 1)]
    [Arguments(-10, 3, -3)]
    [Arguments(-10, -4, 2)]
    [Arguments(1E+100, 1E+40, 1E+60)]
    public void Quotient(double x, double y, double expected)
    {
        ClassicAssert.AreEqual(expected, (double)XLWorkbook.EvaluateExpr($"QUOTIENT({x}, {y})"));
    }

    [Test]
    public void QuotientErrors() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("QUOTIENT(1, 0)"));

    [Test]
    [Arguments(270, 4.71238898038469)]
    [Arguments(-180, -Math.PI)]
    public void Radians(double angle, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"RADIANS({angle})"),
            XLHelper.Epsilon
        );
    }

    [Test]
    public void Rand()
    {
        for (int i = 0; i < 100; ++i)
        {
            double randomNumber = (double)XLWorkbook.EvaluateExpr("RAND()");
            ClassicAssert.IsTrue(randomNumber >= 0 && randomNumber < 1);
        }
    }

    [Test]
    public void RandBetween()
    {
        for (int i = 0; i < 100; ++i)
        {
            double randomNumber = (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(10, 20)");
            ClassicAssert.IsTrue(randomNumber >= 10 && randomNumber <= 20);
        }

        ClassicAssert.AreEqual(101, (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(100.5, 100.9)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("RANDBETWEEN(100.9, 100.5)")
        );
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("RANDBETWEEN(20, 5)")
        );
        double randBetween = (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(1E+100, 1E+110)");
        ClassicAssert.IsTrue(randBetween >= 1E+100 && randBetween <= 1E+110);
    }

    [Test]
    [Arguments(1, 0, "I")]
    [Arguments(3046, 1, @"MMMVLI")]
    [Arguments(3999, 1, @"MMMLMVLIV")]
    [Arguments(999, 0, @"CMXCIX")]
    [Arguments(999.99, 0.9, @"CMXCIX")]
    [Arguments(999, 1, @"LMVLIV")]
    [Arguments(999, 2, @"XMIX")]
    [Arguments(999, 3, @"VMIV")]
    [Arguments(999, 4, @"IM")]
    public void Roman(double value, double form, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (string)XLWorkbook.EvaluateExpr($"ROMAN({value}, {form})")
        );
    }

    [Test]
    public void RomanValue0IsEmptyString() =>
        ClassicAssert.AreEqual(string.Empty, XLWorkbook.EvaluateExpr("ROMAN(0, 0)"));

    [Test]
    public void RomanHasOptionalSecondArgumentWithDefaultValue0() =>
        ClassicAssert.AreEqual(@"CMXCIX", XLWorkbook.EvaluateExpr("ROMAN(999)"));

    [Test]
    public void RomanFormMustBeBetween0And4()
    {
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(1, -1)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(1, 5)"));
    }

    [Test]
    public void RomanValueMustBeBetween0And3999()
    {
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(-1, 0)"));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("ROMAN(4000, 0)")
        );
    }

    [Test]
    [Arguments(2.15, 1, 2.2)]
    [Arguments(2.149, 1, 2.1)]
    [Arguments(-1.475, 2, -1.48)]
    [Arguments(21.5, -1, 20.0)]
    [Arguments(626.3, -3, 1000.0)]
    [Arguments(1.98, -1, 0.0)]
    [Arguments(-50.55, -2, -100.0)]
    [Arguments(31.565, 2, 31.57)]
    [Arguments(-31.565, 2, -31.57)]
    [Arguments(1E+100, 2, 1E+100)]
    [Arguments(1.25, 0, 1)]
    [Arguments(1, -1E+100, 0)]
    [Arguments(1.123456, 1E+100, 1.123456)] // Excel says 0 for anything over 2147483646
    public void Round(double number, double digits, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"ROUND({number}, {digits})")
        );
    }

    [Test]
    [Arguments(3.2, 0, 3.0)]
    [Arguments(76.9, 0, 76.0)]
    [Arguments(3.14159, 3, 3.141)]
    [Arguments(-3.14159, 1, -3.1)]
    [Arguments(31415.92654, -2, 31400.0)]
    [Arguments(0, 3, 0)]
    public void RoundDown(double number, double digits, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"ROUNDDOWN({number}, {digits})")
        );
    }

    [Test]
    [Arguments(3.2, 0, 4)]
    [Arguments(76.9, 0, 77.0)]
    [Arguments(3.14159, 3, 3.142)]
    [Arguments(-3.14159, 1, -3.2)]
    [Arguments(31415.92654, -2, 31500.0)]
    [Arguments(0, 3, 0)]
    [Arguments(11, 0, 11)]
    public void RoundUp(double number, double digits, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"ROUNDUP({number}, {digits})")
        );
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("10.5", 1)]
    [Arguments("-5.4", -1)]
    [Arguments("-0.00001", -1)]
    [Arguments("-1E+300", -1)]
    [Arguments("\"0 1/2\"", 1)]
    [Arguments("FALSE", 0)]
    [Arguments("TRUE", 1)]
    public void Sign(string arg, double expected)
    {
        ClassicAssert.AreEqual(expected, (double)XLWorkbook.EvaluateExpr($"SIGN({arg})"));
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("1", 0.8414709848078965)]
    [Arguments("-1", -0.8414709848078965)]
    [Arguments("PI()", 0)]
    [Arguments("PI()/2", 1)]
    [Arguments("30*PI()/180", 0.5)]
    [Arguments("RADIANS(30)", 0.5)]
    public void Sin(string arg, double expected)
    {
        ClassicAssert.AreEqual(expected, (double)XLWorkbook.EvaluateExpr($"SIN({arg})"), Tolerance);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("1", 1.1752011936438014)]
    [Arguments("10", 11013.232874703393)]
    [Arguments("100", 1.3440585709080678E+43)]
    [Arguments("100", 1.3440585709080678E+43)]
    [Arguments("711", XLError.NumberInvalid)]
    [Arguments("-711", XLError.NumberInvalid)]
    public void Sinh(string arg, object result)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"SINH({arg})");
        ClassicAssert.AreEqual(result, actual);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(0.3, 1.0467516)]
    [Arguments(0.6, 1.21162831)]
    [Arguments(0.9, 1.60872581)]
    [Arguments(1.2, 2.759703601)]
    [Arguments(1.5, 14.1368329)]
    [Arguments(1.8, -4.401367872)]
    [Arguments(2.1, -1.980801656)]
    [Arguments(2.4, -1.356127641)]
    [Arguments(2.7, -1.10610642)]
    [Arguments(3.0, -1.010108666)]
    [Arguments(3.3, -1.012678974)]
    [Arguments(3.6, -1.115127532)]
    [Arguments(3.9, -1.377538917)]
    [Arguments(4.2, -2.039730601)]
    [Arguments(4.5, -4.743927548)]
    [Arguments(4.8, 11.42870421)]
    [Arguments(5.1, 2.645658426)]
    [Arguments(5.4, 1.575565187)]
    [Arguments(5.7, 1.198016873)]
    [Arguments(6.0, 1.041481927)]
    [Arguments(6.3, 1.000141384)]
    [Arguments(6.6, 1.052373922)]
    [Arguments(6.9, 1.225903187)]
    [Arguments(7.2, 1.643787029)]
    [Arguments(7.5, 2.884876262)]
    [Arguments(7.8, 18.53381902)]
    [Arguments(8.1, -4.106031636)]
    [Arguments(8.4, -1.925711244)]
    [Arguments(8.7, -1.335743646)]
    [Arguments(9.0, -1.097537906)]
    [Arguments(9.3, -1.007835594)]
    [Arguments(9.6, -1.015550252)]
    [Arguments(9.9, -1.124617578)]
    [Arguments(10.2, -1.400039323)]
    [Arguments(10.5, -2.102886109)]
    [Arguments(10.8, -5.145888341)]
    [Arguments(11.1, 9.593612018)]
    [Arguments(11.4, 2.541355049)]
    [Arguments(45, 1.90359)]
    [Arguments(30, 6.48292)]
    public void SecReturnsCorrectNumber(double angle, double expectedOutput)
    {
        double result = (double)XLWorkbook.EvaluateExpr($"SEC({angle})");
        ClassicAssert.AreEqual(expectedOutput, result, 0.00001);

        // as the secant is symmetric for positive and negative numbers, let's assert twice:
        double resultForNegative = (double)XLWorkbook.EvaluateExpr($"SEC({-angle})");
        ClassicAssert.AreEqual(expectedOutput, resultForNegative, 0.00001);
    }

    [Test]
    [Arguments(-9, 0.00024682)]
    [Arguments(-8, 0.000670925)]
    [Arguments(-7, 0.001823762)]
    [Arguments(-6, 0.004957474)]
    [Arguments(-5, 0.013475282)]
    [Arguments(-4, 0.036618993)]
    [Arguments(-3, 0.099327927)]
    [Arguments(-2, 0.265802229)]
    [Arguments(-1, 0.648054274)]
    [Arguments(0, 1)]
    [Arguments(1E+100, 0)]
    [Arguments(1E-100, 1)]
    public void SechReturnsCorrectNumber(double angle, double expectedOutput)
    {
        double result = (double)XLWorkbook.EvaluateExpr($"SECH({angle})");
        ClassicAssert.AreEqual(expectedOutput, result, 0.00001);

        // as the secant is symmetric for positive and negative numbers, let's assert twice:
        double resultForNegative = (double)XLWorkbook.EvaluateExpr($"SECH({-angle})");
        ClassicAssert.AreEqual(expectedOutput, resultForNegative, 0.00001);
    }

    [Test]
    public void SeriesSum()
    {
        ClassicAssert.AreEqual(
            40.0,
            (double)XLWorkbook.EvaluateExpr("SERIESSUM(2,3,4,5)"),
            Tolerance
        );

        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A2").FormulaA1 = "PI()/4";
        ws.Cell("A3").Value = 1;
        ws.Cell("A4").FormulaA1 = "-1/FACT(2)";
        ws.Cell("A5").FormulaA1 = "1/FACT(4)";
        ws.Cell("A6").FormulaA1 = "-1/FACT(6)";

        XLCellValue actual = ws.Evaluate("SERIESSUM(A2,0,2,A3:A6)");
        ClassicAssert.AreEqual(0.70710321482284566, (double)actual, Tolerance);
    }

    [Test]
    [Arguments("{1,2,3;4,5,6}")]
    [Arguments("{1,2,3,4,5,6}")]
    [Arguments("{1,2;3,4;5,6}")]
    public void SeriesSumTakesCoefficientsRowByRowLeftToRight(string array) =>
        ClassicAssert.AreEqual(1284, XLWorkbook.EvaluateExpr($"SERIESSUM(2,2,1,{array})"));

    [Test]
    public void SeriesSumReturnsInvalidNumberErrorWhenResultIsTooLarge()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[] { 1, 2, 3, 4, 5 });
        ClassicAssert.AreEqual(3E+300, ws.Evaluate("SERIESSUM(10,100,100,A1:A3)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, ws.Evaluate("SERIESSUM(10,100,100,A1:A4)"));
    }

    [Test]
    public void SeriesSumCoercion()
    {
        // For some weird reason, SERIESSUM doesn't convert logical
        foreach (string invalidValue in new[] { "\"\"", "TRUE" })
        {
            ClassicAssert.AreEqual(
                XLError.IncompatibleValue,
                XLWorkbook.EvaluateExpr($"SERIESSUM({invalidValue},1,1,1)")
            );
            ClassicAssert.AreEqual(
                XLError.IncompatibleValue,
                XLWorkbook.EvaluateExpr($"SERIESSUM(1,{invalidValue},1,1)")
            );
            ClassicAssert.AreEqual(
                XLError.IncompatibleValue,
                XLWorkbook.EvaluateExpr($"SERIESSUM(1,1,{invalidValue},1)")
            );
            ClassicAssert.AreEqual(
                XLError.IncompatibleValue,
                XLWorkbook.EvaluateExpr($"SERIESSUM(1,1,1,{invalidValue})")
            );
        }

        // Blank and text values are coerced to a number
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        foreach (string validArg in new[] { "A1", "\"0 0/2\"" })
        {
            ClassicAssert.AreEqual(0, ws.Evaluate($"SERIESSUM({validArg},1,1,1)"));
            ClassicAssert.AreEqual(1, ws.Evaluate($"SERIESSUM(1,{validArg},1,1)"));
            ClassicAssert.AreEqual(1, ws.Evaluate($"SERIESSUM(1,1,{validArg},1)"));
        }

        // Text is not converted in an area and causes conversion error
        ws.Cell("B2").Value = "0";
        ws.Cell("B3").Value = 5;
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SERIESSUM(1,1,1,B2:B3)"));

        // Blank is interpreted as 0
        ws.Cell("C1").Value = Blank.Value;
        ws.Cell("C2").Value = 2;
        ClassicAssert.AreEqual(2, ws.Evaluate("SERIESSUM(1,1,1,C1:C2)"));
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(1, 1)]
    [Arguments(2, 1.4142135624)]
    [Arguments(1E+300, 1E+150)]
    public void Sqrt(double x, double result) =>
        ClassicAssert.AreEqual(result, (double)XLWorkbook.EvaluateExpr($"SQRT({x})"), Tolerance);

    [Test]
    [Arguments(-1)]
    [Arguments(-0.0001)]
    public void SqrtReturnsInvalidNumberForNegativeNumbers(double x) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"SQRT({x})"));

    [Test]
    public void SqrtPi()
    {
        double actual = (double)XLWorkbook.EvaluateExpr("SQRTPI(1)");
        ClassicAssert.AreEqual(1.7724538509055159, actual, Tolerance);

        actual = (double)XLWorkbook.EvaluateExpr("SQRTPI(2)");
        ClassicAssert.AreEqual(2.5066282746310002, actual, Tolerance);

        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("SQRTPI(-1)"));
    }

    [Test]
    public void Subtotal()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Non-existent functions return error
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(0, A1)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(0.9, A1)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(12, A1)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(100.9, A1)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(112, A1)"));
    }

    [Test]
    public void SubtotalAverage()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").FormulaA1 = "SUBTOTAL(1,A1,A2)";
        ws.Cell("A4").Value = "A";

        ClassicAssert.AreEqual(2.5, ws.Cell("A3").Value);
        ClassicAssert.AreEqual(2.5, ws.Evaluate("SUBTOTAL(1, A1:A4)"));

        ws.Row(2).Hide();
        ClassicAssert.AreEqual(2, ws.Evaluate("SUBTOTAL(101, A1:A4)"));
    }

    [Test]
    public void Subtotal10Calc()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.DefinedNames.Add("subtotalrange", "$A$37:$A$38");

        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("A3").FormulaA1 = "SUBTOTAL(9, A1:A2)"; // simple add subtotal
        ws.Cell("A4").Value = 8;
        ws.Cell("A5").Value = 16;
        ws.Cell("A6").FormulaA1 = "SUBTOTAL(9, A4:A5)"; // simple add subtotal
        ws.Cell("A7").Value = 32;
        ws.Cell("A8").Value = 64;
        ws.Cell("A9").FormulaA1 = "SUM(A7:A8)"; // func but not subtotal
        ws.Cell("A10").Value = 128;
        ws.Cell("A11").Value = 256;
        ws.Cell("A12").FormulaA1 = "SUBTOTAL(1, A10:A11)"; // simple avg subtotal
        ws.Cell("A13").Value = 512;
        ws.Cell("A14").FormulaA1 = "SUBTOTAL(9, A1:A13)"; // subtotals in range
        ws.Cell("A15").Value = 1024;
        ws.Cell("A16").Value = 2048;
        ws.Cell("A17").FormulaA1 = "42 + SUBTOTAL(9, A15:A16)"; // simple add subtotal in formula
        ws.Cell("A18").Value = 4096;
        ws.Cell("A19").FormulaA1 = "SUBTOTAL(9, A15:A18)"; // subtotals in range
        ws.Cell("A20").Value = 8192;
        ws.Cell("A21").Value = 16384;
        ws.Cell("A22").FormulaA1 = @"32768 * SEARCH(""SUBTOTAL(9, A1:A2)"", A28)"; // subtotal literal in formula
        ws.Cell("A23").FormulaA1 = "SUBTOTAL(9, A20:A22)"; // subtotal literal in formula in range
        ws.Cell("A24").Value = 65536;
        ws.Cell("A25").FormulaA1 = "A23"; // link to subtotal
        ws.Cell("A26").FormulaA1 = "PRODUCT(SUBTOTAL(9, A24:A25), 2)"; // subtotal as parameter in func
        ws.Cell("A27").Value = 131072;
        ws.Cell("A28").Value = "SUBTOTAL(9, A1:A2)"; // subtotal literal
        ws.Cell("A29").FormulaA1 = "SUBTOTAL(9, A27:A28)"; // subtotal literal in range
        ws.Cell("A30").FormulaA1 = "SUBTOTAL(9, A31:A32)"; // simple add subtotal backward
        ws.Cell("A31").Value = 262144;
        ws.Cell("A32").Value = 524288;
        ws.Cell("A33").FormulaA1 = "SUBTOTAL(9, A20:A32)"; // subtotals in range
        ws.Cell("A34").FormulaA1 = @"SUBTOTAL(VALUE(""9""), A1:A33, A35:A41)"; // func as parameter in subtotal and many ranges
        ws.Cell("A35").Value = 1048576;
        ws.Cell("A36").FormulaA1 = "SUBTOTAL(9, A31:A32, A35)"; // many ranges
        ws.Cell("A37").Value = 2097152;
        ws.Cell("A38").Value = 4194304;
        ws.Cell("A39").FormulaA1 = "SUBTOTAL(3*3, subtotalrange)"; // formula as parameter in subtotal and named range
        ws.Cell("A40").Value = 8388608;
        ws.Cell("A41").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2)"; // formula with link as parameter in subtotal
        ws.Cell("A42").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2) + SUBTOTAL(A4+1, A35:A40)"; // two subtotals in one formula

        ClassicAssert.AreEqual(6, ws.Cell("A3").Value);
        ClassicAssert.AreEqual(24, ws.Cell("A6").Value);
        ClassicAssert.AreEqual(192, ws.Cell("A12").Value);
        ClassicAssert.AreEqual(1118, ws.Cell("A14").Value);
        ClassicAssert.AreEqual(3114, ws.Cell("A17").Value);
        ClassicAssert.AreEqual(7168, ws.Cell("A19").Value);
        ClassicAssert.AreEqual(57344, ws.Cell("A23").Value);
        ClassicAssert.AreEqual(245760, ws.Cell("A26").Value);
        ClassicAssert.AreEqual(131072, ws.Cell("A29").Value);
        ClassicAssert.AreEqual(786432, ws.Cell("A30").Value);
        ClassicAssert.AreEqual(1097728, ws.Cell("A33").Value);
        ClassicAssert.AreEqual(16834654, ws.Cell("A34").Value);
        ClassicAssert.AreEqual(1835008, ws.Cell("A36").Value);
        ClassicAssert.AreEqual(6291456, ws.Cell("A39").Value);
        ClassicAssert.AreEqual(31457280, ws.Cell("A41").Value);
        ClassicAssert.AreEqual(47185920, ws.Cell("A42").Value);
    }

    [Test]
    public void Subtotal100Calc()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("C1").Value = Blank.Value;
        ws.Cell("A2").Value = "A";
        ws.Cell("B2").Value = 4;
        ws.Cell("C2").Value = 8;
        ws.Cell("A3").FormulaA1 = "SUBTOTAL(109, A1:A2)";
        ws.Cell("B3").FormulaA1 = "SUBTOTAL(109, B1:B2)";
        ws.Cell("C3").FormulaA1 = "SUBTOTAL(109, C1:C2)";
        ws.Cell("A4").Value = 16;
        ws.Cell("B4").Value = 32;
        ws.Cell("C4").Value = 64;
        ws.Cell("A5").Value = 128;
        ws.Cell("B5").Value = 256;
        ws.Cell("C5").Value = 512;
        ws.Cell("A6").FormulaA1 = "SUBTOTAL(109, A1:A5)";
        ws.Cell("B6").FormulaA1 = "SUBTOTAL(109, B1:B5)";
        ws.Cell("C6").FormulaA1 = "SUBTOTAL(109, C1:C5)";

        ws.Row(2).Hide();
        ws.Row(5).Hide();

        ClassicAssert.AreEqual(1, ws.Cell("A3").Value);
        ClassicAssert.AreEqual(2, ws.Cell("B3").Value);
        ClassicAssert.AreEqual(0, ws.Cell("C3").Value);
        ClassicAssert.AreEqual(17, ws.Cell("A6").Value);
        ClassicAssert.AreEqual(34, ws.Cell("B6").Value);
        ClassicAssert.AreEqual(64, ws.Cell("C6").Value);
    }

    [Test]
    public void SubtotalCount()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(2,A1:A3)";

        ClassicAssert.AreEqual(2, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(1, ws.Evaluate("SUBTOTAL(2,A2:A4)"));

        ws.Row(2).Hide();
        ClassicAssert.AreEqual(1, ws.Evaluate("SUBTOTAL(102,A1:A4)"));
    }

    [Test]
    public void SubtotalCountA()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = string.Empty;
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(3,A1,A2,A3)";

        ClassicAssert.AreEqual(3, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(3, ws.Evaluate("SUBTOTAL(3,A1:A4)"));

        ws.Row(1).Hide();
        ClassicAssert.AreEqual(2, ws.Evaluate("SUBTOTAL(103,A1:A4)"));
    }

    [Test]
    public void SubtotalMax()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(4,A1,A2,A3) + 10";

        ClassicAssert.AreEqual(13, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(3, ws.Evaluate("SUBTOTAL(4,A1:A4)"));

        ws.Cell("A5").Value = 2.5;
        ws.Row(2).Hide();
        ClassicAssert.AreEqual(2.5, ws.Evaluate("SUBTOTAL(104,A1:A5)"));
    }

    [Test]
    public void SubtotalMin()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(5,A1,A2,A3) - 10";

        ClassicAssert.AreEqual(-8, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(2, ws.Evaluate("SUBTOTAL(5,A1:A4)"));

        ws.Cell("A5").Value = 2.5;
        ws.Row(1).Hide();
        ClassicAssert.AreEqual(2.5, ws.Evaluate("SUBTOTAL(105,A1:A5)"));
    }

    [Test]
    public void SubtotalProduct()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(6,A1,A2,A3)";

        ClassicAssert.AreEqual(6, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(6, ws.Evaluate("SUBTOTAL(6,A1:A4)"));

        ws.Row(2).Hide();
        ws.Cell("A5").Value = 4;
        ClassicAssert.AreEqual(8, ws.Evaluate("SUBTOTAL(106,A1:A5)"));
    }

    [Test]
    public void SubtotalStDev()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(7,A1:A3,A5)";
        ws.Cell("A5").Value = 5;

        ClassicAssert.AreEqual(1.5275252316, (double)ws.Cell("A4").Value, XLHelper.Epsilon);
        ClassicAssert.AreEqual(
            1.5275252316,
            (double)ws.Evaluate("SUBTOTAL(7,A1:A5)"),
            XLHelper.Epsilon
        );

        ws.Row(2).Hide();
        ClassicAssert.AreEqual(
            2.1213203435,
            (double)ws.Evaluate("SUBTOTAL(107,A1:A5)"),
            XLHelper.Epsilon
        );
    }

    [Test]
    public void SubtotalStDevP()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(8,A1,A2,A3)";

        ClassicAssert.AreEqual(0.5, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(0.5, ws.Evaluate("SUBTOTAL(8,A1:A4)"));

        ws.Row(2).Hide();
        ws.Cell("A5").Value = 3;
        ClassicAssert.AreEqual(0.5, ws.Evaluate("SUBTOTAL(108,A1:A5)"));
    }

    [Test]
    public void SubtotalSum()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(9,A1,A2,A3)";

        ClassicAssert.AreEqual(5, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(5, ws.Evaluate("SUBTOTAL(9,A1:A4)"));

        ws.Row(2).Hide();

        ClassicAssert.AreEqual(2, ws.Evaluate("SUBTOTAL(109, A1:A4)"));
    }

    [Test]
    public void SubtotalVar()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        ws.Cell("A2").Value = 4;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").Value = 8;
        ws.Cell("A5").Value = 5;
        ws.Cell("A6").FormulaA1 = "SUBTOTAL(10,A1:A5)";

        ClassicAssert.AreEqual(3, ws.Cell("A6").Value);
        ClassicAssert.AreEqual(3, ws.Evaluate("SUBTOTAL(10,A1:A6)"));

        ws.Row(1).Hide();
        ws.Row(5).Hide();
        ClassicAssert.AreEqual(8, ws.Evaluate("SUBTOTAL(110,A1:A6)"));
    }

    [Test]
    public void SubtotalVarP()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(11,A1,A2,A3)";

        ClassicAssert.AreEqual(0.25, ws.Cell("A4").Value);
        ClassicAssert.AreEqual(0.25, ws.Evaluate("SUBTOTAL(11,A1:A4)"));

        ws.Row(2).Hide();
        ws.Cell("A5").Value = 4;
        ClassicAssert.AreEqual(1, ws.Evaluate("SUBTOTAL(111,A1:A5)"));
    }

    [Test]
    public void Sum()
    {
        IXLCell cell = new XLWorkbook().AddWorksheet("Sheet1").FirstCell();
        IXLCell fCell = cell.SetValue(1).CellBelow().SetValue(2).CellBelow();
        fCell.FormulaA1 = "sum(A1:A2)";

        ClassicAssert.AreEqual(3.0, fCell.Value);
    }

    [Test]
    public void SumDateTimeAndNumber()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = new DateTime(2018, 1, 1);
            ClassicAssert.AreEqual(43102, ws.Evaluate("SUM(A1:A2)"));

            ws.Cell("A1").Value = 2;
            ws.Cell("A2").FormulaA1 = "DATE(2018,1,1)";
            ClassicAssert.AreEqual(43103, ws.Evaluate("SUM(A1:A2)"));
        }
    }

    [Test]
    [Arguments(9, "SUMIF(A:B, \"A*\", C:C)")]
    [Arguments(9, "SUMIF(A1:B6, \"A*\", C1:C6)")]
    public void SumIfInputRangeHasMultipleColumns(int expectedOutcome, string formula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Data");
        object[] data =
        [
            new
            {
                Id = "AA",
                Id2 = "BA",
                Value = 2,
            },
            new
            {
                Id = "AB",
                Id2 = "BB",
                Value = 3,
            },
            new
            {
                Id = "BA",
                Id2 = "AA",
                Value = 2,
            },
            new
            {
                Id = "BB",
                Id2 = "AB",
                Value = 1,
            },
            new
            {
                Id = "AC",
                Id2 = "AC",
                Value = 4,
            },
        ];
        ws.Cell("A1").InsertTable(data);

        ClassicAssert.AreEqual(expectedOutcome, ws.Evaluate(formula));
    }

    /// <summary>
    /// refers to Example 1 from the Excel documentation,
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedOutcome"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(63000, "SUMIF(A1:A4,\">160000\", B1:B4)")]
    [Arguments(900000, "SUMIF(A1:A4,\">160000\")")]
    [Arguments(21000, "SUMIF(A1:A4, 300000, B1:B4)")]
    [Arguments(28000, "SUMIF(A1:A4, \">\" &C1, B1:B4)")]
    public void SumIfReturnsCorrectValuesReferenceExample1FromMicrosoft(
        int expectedOutcome,
        string formula
    )
    {
        using (XLWorkbook wb = new())
        {
            wb.ReferenceStyle = XLReferenceStyle.A1;

            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell(1, 1).Value = 100000;
            ws.Cell(1, 2).Value = 7000;
            ws.Cell(2, 1).Value = 200000;
            ws.Cell(2, 2).Value = 14000;
            ws.Cell(3, 1).Value = 300000;
            ws.Cell(3, 2).Value = 21000;
            ws.Cell(4, 1).Value = 400000;
            ws.Cell(4, 2).Value = 28000;

            ws.Cell(1, 3).Value = 300000;

            ClassicAssert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
        }
    }

    /// <summary>
    /// refers to Example 2 from the Excel documentation,
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedOutcome"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(2000, "SUMIF(A2:A7,\"Fruits\", C2:C7)")]
    [Arguments(12000, "SUMIF(A2:A7,\"Vegetables\", C2:C7)")]
    [Arguments(4300, "SUMIF(B2:B7, \"*es\", C2:C7)")]
    [Arguments(400, "SUMIF(A2:A7, \"\", C2:C7)")]
    public void SumIfReturnsCorrectValuesReferenceExample2FromMicrosoft(
        int expectedOutcome,
        string formula
    )
    {
        using (XLWorkbook wb = new())
        {
            wb.ReferenceStyle = XLReferenceStyle.A1;

            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell(2, 1).Value = "Vegetables";
            ws.Cell(3, 1).Value = "Vegetables";
            ws.Cell(4, 1).Value = "Fruits";
            ws.Cell(5, 1).Value = "";
            ws.Cell(6, 1).Value = "Vegetables";
            ws.Cell(7, 1).Value = "Fruits";

            ws.Cell(2, 2).Value = "Tomatoes";
            ws.Cell(3, 2).Value = "Celery";
            ws.Cell(4, 2).Value = "Oranges";
            ws.Cell(5, 2).Value = "Butter";
            ws.Cell(6, 2).Value = "Carrots";
            ws.Cell(7, 2).Value = "Apples";

            ws.Cell(2, 3).Value = 2300;
            ws.Cell(3, 3).Value = 5500;
            ws.Cell(4, 3).Value = 800;
            ws.Cell(5, 3).Value = 400;
            ws.Cell(6, 3).Value = 4200;
            ws.Cell(7, 3).Value = 1200;

            ws.Cell(1, 3).Value = 300000;

            ClassicAssert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
        }
    }

    [Test]
    public void SumIfReturnsCorrectValuesWhenCalledOnFullColumn()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Data");
            object[] data =
            [
                new { Id = "A", Value = 2 },
                new { Id = "B", Value = 3 },
                new { Id = "C", Value = 2 },
                new { Id = "A", Value = 1 },
                new { Id = "B", Value = 4 },
            ];
            ws.Cell("A1").InsertTable(data);
            string formula = "=SUMIF(A:A,\"=A\",B:B)";
            XLCellValue value = ws.Evaluate(formula);
            ClassicAssert.AreEqual(3, value);
        }
    }

    [Test]
    public void SumIfReturnsCorrectValuesWhenFormulaBelongToSameRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Data");
            object[] data =
            [
                new { Id = "A", Value = 2 },
                new { Id = "B", Value = 3 },
                new { Id = "C", Value = 2 },
                new { Id = "A", Value = 1 },
                new { Id = "B", Value = 4 },
            ];
            ws.Cell("A1").InsertTable(data);
            ws.Cell("A7").SetValue("Sum A");
            // SUMIF formula
            string formula = "=SUMIF(A:A,\"=A\",B:B)";
            ws.Cell("B7").SetFormulaA1(formula);
            XLCellValue value = ws.Cell("B7").Value;
            ClassicAssert.AreEqual(3, value);
        }
    }

    [Test]
    public void SumIfsMultidimensionalRanges()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell()
            .InsertData(
                new object[]
                {
                    (10, 10, 1, 2),
                    (20, 15, 2, 4),
                    (30, 20, 3, 6),
                    (40, 25, 4, 8),
                    (50, 30, 5, 10),
                }
            );
        ClassicAssert.AreEqual(30, ws.Evaluate("SUMIFS(C1:D5,A1:B5,\">20\")"));
    }

    /// <summary>
    /// refers to Example 2 to SumIf from the Excel documentation.
    /// As SumIfs should behave the same if called with three parameters, we can take that example here again.
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedResult"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(2000, "SUMIFS(C2:C7, A2:A7, \"Fruits\")")]
    [Arguments(12000, "SUMIFS(C2:C7, A2:A7, \"Vegetables\")")]
    [Arguments(4300, "SUMIFS(C2:C7, B2:B7, \"*es\")")]
    [Arguments(400, "SUMIFS(C2:C7, A2:A7, \"\")")]
    public void SumIfsReturnsCorrectValuesReferenceExample2FromMicrosoft(
        int expectedResult,
        string formula
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell(2, 1).Value = "Vegetables";
            ws.Cell(3, 1).Value = "Vegetables";
            ws.Cell(4, 1).Value = "Fruits";
            ws.Cell(5, 1).Value = "";
            ws.Cell(6, 1).Value = "Vegetables";
            ws.Cell(7, 1).Value = "Fruits";

            ws.Cell(2, 2).Value = "Tomatoes";
            ws.Cell(3, 2).Value = "Celery";
            ws.Cell(4, 2).Value = "Oranges";
            ws.Cell(5, 2).Value = "Butter";
            ws.Cell(6, 2).Value = "Carrots";
            ws.Cell(7, 2).Value = "Apples";

            ws.Cell(2, 3).Value = 2300;
            ws.Cell(3, 3).Value = 5500;
            ws.Cell(4, 3).Value = 800;
            ws.Cell(5, 3).Value = 400;
            ws.Cell(6, 3).Value = 4200;
            ws.Cell(7, 3).Value = 1200;

            ws.Cell(1, 3).Value = 300000;

            double actualResult = (double)ws.Evaluate(formula);
            ClassicAssert.AreEqual(expectedResult, actualResult);
        }
    }

    /// <summary>
    /// refers to Example 1 to SumIf from the Excel documentation.
    /// As SumIfs should behave the same if called with three parameters, but in a different order
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedOutcome"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(63000, "SUMIFS(B1:B4, A1:A4, \">160000\")")]
    [Arguments(21000, "SUMIFS(B1:B4, A1:A4, 300000)")]
    [Arguments(28000, "SUMIFS(B1:B4, A1:A4, \">\" &C1)")]
    public void SumIfsReturnsCorrectValuesReferenceExampleForSumIf1FromMicrosoft(
        int expectedOutcome,
        string formula
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell(1, 1).Value = 100000;
            ws.Cell(1, 2).Value = 7000;
            ws.Cell(2, 1).Value = 200000;
            ws.Cell(2, 2).Value = 14000;
            ws.Cell(3, 1).Value = 300000;
            ws.Cell(3, 2).Value = 21000;
            ws.Cell(4, 1).Value = 400000;
            ws.Cell(4, 2).Value = 28000;

            ws.Cell(1, 3).Value = 300000;

            ClassicAssert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
        }
    }

    /// <summary>
    /// refers to example data and formula to SumIfs in the Excel documentation,
    /// <see cref="https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    [Test]
    [Arguments(20, "=SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")")]
    [Arguments(30, "=SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")")]
    public void SumIfsReturnsCorrectValuesReferenceExampleFromMicrosoft(
        int expectedResult,
        string formula
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            int row = 2;

            ws.Cell(row, 1).Value = 5;
            ws.Cell(row, 2).Value = "Apples";
            ws.Cell(row, 3).Value = "Tom";
            row++;

            ws.Cell(row, 1).Value = 4;
            ws.Cell(row, 2).Value = "Apples";
            ws.Cell(row, 3).Value = "Sarah";
            row++;

            ws.Cell(row, 1).Value = 15;
            ws.Cell(row, 2).Value = "Artichokes";
            ws.Cell(row, 3).Value = "Tom";
            row++;

            ws.Cell(row, 1).Value = 3;
            ws.Cell(row, 2).Value = "Artichokes";
            ws.Cell(row, 3).Value = "Sarah";
            row++;

            ws.Cell(row, 1).Value = 22;
            ws.Cell(row, 2).Value = "Bananas";
            ws.Cell(row, 3).Value = "Tom";
            row++;

            ws.Cell(row, 1).Value = 12;
            ws.Cell(row, 2).Value = "Bananas";
            ws.Cell(row, 3).Value = "Sarah";
            row++;

            ws.Cell(row, 1).Value = 10;
            ws.Cell(row, 2).Value = "Carrots";
            ws.Cell(row, 3).Value = "Tom";
            row++;

            ws.Cell(row, 1).Value = 33;
            ws.Cell(row, 2).Value = "Carrots";
            ws.Cell(row, 3).Value = "Sarah";

            XLCellValue actualResult = ws.Evaluate(formula);

            ClassicAssert.AreEqual(expectedResult, (double)actualResult, Tolerance);
        }
    }

    [Test]
    [Arguments("SUMIFS(D1:E5,A1:B5,\"A*\",C1:C5,\">2\")")]
    [Arguments("SUMIFS(H1:I3,A1:B3,1,D1:F2,2)")]
    [Arguments("SUMIFS(D:E,A:B,\"A*\",C:C,\">2\")")]
    public void SumIfsReturnsErrorWhenRangeDimensionsAreNotSame(string formula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula));
    }

    [Test]
    [Arguments("SUMIFS(A1:A3, B1:B3,\"<>B\")", 11)]
    [Arguments("SUMIFS(A1:A3, B1:B3,\"<>\")", 110)]
    public void SumIfsMatchesBlankCellsWhenCriteriaIsNotEqual(string formula, double expectedSum)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 10;
        ws.Cell("A3").Value = 100;
        ws.Cell("B1").Value = Blank.Value;
        ws.Cell("B2").Value = string.Empty;
        ws.Cell("B3").Value = "B";

        ClassicAssert.AreEqual(expectedSum, ws.Evaluate(formula));
    }

    [Test]
    public void SumProduct()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");

        ws.FirstCell().InsertData(Enumerable.Range(1, 10));
        ws.FirstCell().CellRight().InsertData(Enumerable.Range(1, 10).Reverse());

        ClassicAssert.AreEqual(2, ws.Evaluate("SUMPRODUCT(A2)"));
        ClassicAssert.AreEqual(55, ws.Evaluate("SUMPRODUCT(A1:A10)"));
        ClassicAssert.AreEqual(220, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B5)"));

        // Scalar, one element array and single cell area are compatible
        ClassicAssert.AreEqual(60, ws.Evaluate("SUMPRODUCT(A5, 4, {3})"));

        // An array can be an argument
        ClassicAssert.AreEqual(10, ws.Evaluate("SUMPRODUCT(A1:A3, {3;2;1})"));

        // An array must have correct orientation, otherwise dimensions don't match
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            ws.Evaluate("SUMPRODUCT(A1:A3, {3,2,1})")
        );

        // Anything but number is counted as zero. The second array is zero for all values = result is 0.
        ClassicAssert.AreEqual(0, ws.Evaluate("SUMPRODUCT({1,2,3,4}, {TRUE,FALSE,\"1\",\"\"})"));

        // Any error returns error
        ClassicAssert.AreEqual(
            XLError.NoValueAvailable,
            ws.Evaluate("SUMPRODUCT({1,2}, {1,#N/A})")
        );
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMPRODUCT(A1, #N/A)"));
        ws.Cell("A2").Value = XLError.NoValueAvailable;
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMPRODUCT(A2, 5)"));

        // Blank cells and cells with text should be treated as zeros
        ws.Range("A1:A5").Clear();
        ClassicAssert.AreEqual(110, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

        // Non-number values are treated as zero
        ws.Range("A1:A5").SetValue("asdf");
        ClassicAssert.AreEqual(110, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

        // Blank cell is considered as a blank and cause #VALUE! error
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(Z1, 5)"));

        // Blank value will cause #VALUE! error
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(IF(TRUE,,), 5)"));
    }

    [Test]
    public void SumSq()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Examples from specification
        ClassicAssert.AreEqual(4.0, XLWorkbook.EvaluateExpr("SUMSQ(2)"));
        ClassicAssert.AreEqual(19.21, XLWorkbook.EvaluateExpr("SUMSQ(2.5, -3.6)"));
        ClassicAssert.AreEqual(24.97, XLWorkbook.EvaluateExpr("SUMSQ({ 2.5, -3.6}, 2.4)"));

        // Scalar blank is converted to 0
        ClassicAssert.AreEqual(16, XLWorkbook.EvaluateExpr("SUMSQ(IF(TRUE,), 4)"));

        // Scalar logical is converted to number
        ClassicAssert.AreEqual(10, XLWorkbook.EvaluateExpr("SUMSQ(3, TRUE)"));

        // Scalar text is converted to number
        ClassicAssert.AreEqual(25, XLWorkbook.EvaluateExpr("SUMSQ(\"4\", \"3\")"));

        // Scalar text that is not convertible return error
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("SUMSQ(1, \"Hello\")")
        );

        // Array logical arguments are ignored
        ClassicAssert.AreEqual(4, XLWorkbook.EvaluateExpr("SUMSQ({2,TRUE,TRUE,FALSE,FALSE})"));

        // Array text arguments are ignored
        ClassicAssert.AreEqual(20, XLWorkbook.EvaluateExpr("SUMSQ({4, 2, \"hello\", \"10\" })"));

        // Blank, logical and text from reference are ignored
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = 1;
        ws.Cell("A6").Value = 4;
        ClassicAssert.AreEqual(17, ws.Evaluate("SUMSQ(A1:A6)"));

        // Scalar error is propagated
        ClassicAssert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("SUMSQ(1, #NULL!)"));

        // Array error is propagated
        ClassicAssert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("SUMSQ({1, #NULL!})"));

        // Reference error is propagated
        ws.Cell("A1").Value = XLError.NoValueAvailable;
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMSQ(A1)"));
    }

    [Test]
    [Arguments(-1, -1.5574077247)]
    [Arguments(0, 0)]
    [Arguments(1, 1.5574077247)]
    [Arguments(134217727, 3.2584564256)]
    [Arguments(-134217727, -3.2584564256)]
    public void Tan(double radians, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"TAN({radians})"),
            Tolerance
        );
    }

    [Test]
    [Arguments(134217728)]
    [Arguments(-134217728)]
    [Arguments(1E+100)]
    public void TanReturnsInvalidNumberForRadiansOutsideLimit(double radians) =>
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"TAN({radians})"));

    [Test]
    [Arguments(-1, -0.761594156)]
    [Arguments(0, 0)]
    [Arguments(1, 0.761594156)]
    [Arguments(1E+300, 1)]
    [Arguments(-1E+300, -1)]
    public void Tanh(double number, double result) =>
        ClassicAssert.AreEqual(
            result,
            (double)XLWorkbook.EvaluateExpr($"TANH({number})"),
            Tolerance
        );

    [Test]
    [Arguments(27.64799257, null, 27)]
    [Arguments(0, null, 0)]
    [Arguments(0, 0, 0)]
    [Arguments(3.1415926, 0, 3)]
    [Arguments(3.1415926, 1, 3.1)]
    [Arguments(3.1415926, 3, 3.141)]
    [Arguments(3.1415926, 5, 3.14159)]
    [Arguments(-4.3, 0, -4)]
    [Arguments(8.9, null, 8)]
    [Arguments(-8.9, null, -8)]
    [Arguments(0.45, null, 0)]
    public void Trunc(double number, double? digits, object expectedResult)
    {
        string formula = digits is null ? $"TRUNC({number})" : $"TRUNC({number}, {digits})";
        ClassicAssert.AreEqual(expectedResult, (double)XLWorkbook.EvaluateExpr(formula));
    }

    [Test]
    [Arguments(27.64799257, -1, 20)]
    [Arguments(27.64799257, 0, 27)]
    [Arguments(27.64799257, 1, 27.6)]
    [Arguments(27.64799257, 4, 27.6479)]
    public void TruncSpecifyDigits(double input, int digits, double expectedResult)
    {
        double actual = (double)
            XLWorkbook.EvaluateExpr(
                $"TRUNC({input.ToString(CultureInfo.InvariantCulture)}, {digits})"
            );
        ClassicAssert.AreEqual(expectedResult, actual);
    }
}
