using System.Globalization;
using XlsxSharp.ExcelNumberFormat;

namespace XlsxSharp.Tests.ExcelNumberFormat;

public class ExcelNumberFormatTests
{
    private static string? Format(
        object value,
        string formatString,
        CultureInfo culture,
        bool isDate1904 = false
    )
    {
        NumberFormat format = new(formatString);
        if (format.IsValid)
        {
            return format.Format(value, culture, isDate1904);
        }

        return null;
    }

    private static bool IsDateFormatString(string formatString)
    {
        NumberFormat? format = new(formatString);
        return format?.IsDateTimeFormat ?? false;
    }

    [Test]
    public void TestCondition()
    {
        this.Test("Hello", "\"p\"0;\"m\"0;\"z\"0;\"t\"@", "tHello");
        this.Test("Hello", "\"num\"0", "Hello");

        this.Test(-1, "\"p\"0;\"m\"0;\"z\"0;\"t\"@", "m1");
        this.Test(0, "\"p\"0;\"m\"0;\"z\"0;\"t\"@", "z0");
        this.Test(1, "\"p\"0;\"m\"0;\"z\"0;\"t\"@", "p1");

        this.Test(-1, "[<0]\"p\"0;\"m\"0;\"z\"0;\"t\"@", "p1");
        this.Test(0, "[<0]\"p\"0;\"m\"0;\"z\"0;\"t\"@", "z0");
        this.Test(1, "[<0]\"p\"0;\"m\"0;\"z\"0;\"t\"@", "z1");

        this.Test(-1, "\"p\"0;[>0]\"m\"0;\"z\"0;\"t\"@", "-z1");
        this.Test(0, "\"p\"0;[>0]\"m\"0;\"z\"0;\"t\"@", "z0");
        this.Test(1, "\"p\"0;[>0]\"m\"0;\"z\"0;\"t\"@", "p1");

        this.Test(-1, "[<0]\"LT0\";\"ELSE\"", "LT0");
        this.Test(0, "[<0]\"LT0\";\"ELSE\"", "ELSE");
        this.Test(1, "[<0]\"LT0\";\"ELSE\"", "ELSE");

        this.Test(-1, "[<=0]\"LTE0\";\"ELSE\"", "LTE0");
        this.Test(0, "[<=0]\"LTE0\";\"ELSE\"", "LTE0");
        this.Test(1, "[<=0]\"LTE0\";\"ELSE\"", "ELSE");

        this.Test(-1, "[>0]\"GT0\";\"ELSE\"", "ELSE");
        this.Test(0, "[>0]\"GT0\";\"ELSE\"", "ELSE");
        this.Test(1, "[>0]\"GT0\";\"ELSE\"", "GT0");

        this.Test(-1, "[>=0]\"GTE0\";\"ELSE\"", "ELSE");
        this.Test(0, "[>=0]\"GTE0\";\"ELSE\"", "GTE0");
        this.Test(1, "[>=0]\"GTE0\";\"ELSE\"", "GTE0");

        this.Test(-1, "[=0]\"EQ0\";\"ELSE\"", "ELSE");
        this.Test(0, "[=0]\"EQ0\";\"ELSE\"", "EQ0");
        this.Test(1, "[=0]\"EQ0\";\"ELSE\"", "ELSE");

        this.Test(-1, "[<>0]\"NEQ0\";\"ELSE\"", "NEQ0");
        this.Test(0, "[<>0]\"NEQ0\";\"ELSE\"", "ELSE");
        this.Test(1, "[<>0]\"NEQ0\";\"ELSE\"", "NEQ0");
    }

    [Test]
    public void TestFractionAlignmentSuffix()
    {
        this.Test(0, "??/??", " 0/1 ");
        this.Test(1.5, "??/??", " 3/2 ");
        this.Test(3.4, "??/??", "17/5 ");
        this.Test(4.3, "??/??", "43/10");

        this.Test(0, "00/00", "00/01");
        this.Test(1.5, "00/00", "03/02");
        this.Test(3.4, "00/00", "17/05");
        this.Test(4.3, "00/00", "43/10");

        this.Test(0.00, "# ??/\"a\"?\"a\"0\"a\"", "0        a");
        this.Test(0.10, "# ??/\"a\"?\"a\"0\"a\"", "0        a");
        this.Test(0.12, "# ??/\"a\"?\"a\"0\"a\"", "  1/a8a0a");

        this.Test(1.00, "# ??/\"a\"?\"a\"0\"a\"", "1        a");
        this.Test(1.10, "# ??/\"a\"?\"a\"0\"a\"", "1  1/a9a0a");
        this.Test(1.12, "# ??/\"a\"?\"a\"0\"a\"", "1  1/a8a0a");
    }

    [Test]
    public void TestIsDateFormatString()
    {
        ClassicAssert.IsTrue(IsDateFormatString("dd/mm/yyyy"));
        ClassicAssert.IsTrue(IsDateFormatString("dd/mm/yyyy"));
        ClassicAssert.IsTrue(IsDateFormatString("dd-mmm-yy"));
        ClassicAssert.IsTrue(IsDateFormatString("dd-mmmm"));
        ClassicAssert.IsTrue(IsDateFormatString("mmm-yy"));
        ClassicAssert.IsTrue(IsDateFormatString("h:mm AM/PM"));
        ClassicAssert.IsTrue(IsDateFormatString("h:mm:ss AM/PM"));
        ClassicAssert.IsTrue(IsDateFormatString("hh:mm"));
        ClassicAssert.IsTrue(IsDateFormatString("hh:mm:ss"));
        ClassicAssert.IsTrue(IsDateFormatString("dd/mm/yyyy hh:mm"));
        ClassicAssert.IsTrue(IsDateFormatString("mm:ss"));
        ClassicAssert.IsTrue(IsDateFormatString("mm:ss.0"));
        ClassicAssert.IsTrue(IsDateFormatString("[$-809]dd mmmm yyyy"));
        ClassicAssert.IsFalse(IsDateFormatString("#,##0;[Red]-#,##0"));
        ClassicAssert.IsFalse(IsDateFormatString("0_);[Red](0)"));
        ClassicAssert.IsFalse(IsDateFormatString(@"0\h"));
        ClassicAssert.IsFalse(IsDateFormatString("0\"h\""));
        ClassicAssert.IsFalse(IsDateFormatString("0%"));
        ClassicAssert.IsFalse(IsDateFormatString("General"));
        ClassicAssert.IsFalse(
            IsDateFormatString(
                @"_-* #,##0\ _P_t_s_-;\-* #,##0\ _P_t_s_-;_-* "" - ""??\ _P_t_s_-;_-@_- "
            )
        );
    }

    [Test]
    public void TestDate()
    {
        this.Test(new DateTime(2000, 1, 1), "d-mmm-yy", "1-Jan-00");
        this.Test(
            new DateTime(2000, 1, 1, 12, 34, 56),
            "m/d/yyyy\\ h:mm:ss;@",
            "1/1/2000 12:34:56"
        );
        this.Test(new DateTime(2010, 9, 26), "yyyy-MMM-dd", "2010-Sep-26");
        this.Test(new DateTime(2010, 9, 26), "yyyy-MM-dd", "2010-09-26");
        this.Test(new DateTime(2010, 9, 26), "mm/dd/yyyy", "09/26/2010");
        this.Test(new DateTime(2010, 9, 26), "m/d/yy", "9/26/10");
        this.Test(
            new DateTime(2010, 9, 26, 12, 34, 56, 123),
            "m/d/yy hh:mm:ss.000",
            "9/26/10 12:34:56.123"
        );
        this.Test(
            new DateTime(2010, 9, 26, 12, 34, 56, 123),
            "YYYY-MM-DD HH:MM:SS",
            "2010-09-26 12:34:56"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss AM/PM;@",
            "1/1/2020 2:35:55 PM"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss aM/Pm;@",
            "1/1/2020 2:35:55 PM"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss am/PM;@",
            "1/1/2020 2:35:55 PM"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss A/P;@",
            "1/1/2020 2:35:55 P"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss a/P;@",
            "1/1/2020 2:35:55 p"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss A/p;@",
            "1/1/2020 2:35:55 P"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ h:mm:ss;@",
            "1/1/2020 14:35:55"
        );
        this.Test(
            new DateTime(2020, 1, 1, 14, 35, 55),
            "m/d/yyyy\\ hh:mm:ss AM/PM;@",
            "1/1/2020 02:35:55 PM"
        );
        this.Test(
            new DateTime(2020, 1, 1, 16, 5, 6),
            "m/d/yyyy\\ h:m:s AM/PM;@",
            "1/1/2020 4:5:6 PM"
        );
        this.Test(
            new DateTime(2017, 10, 16, 0, 0, 0),
            "dddd, MMMM d, yyyy",
            "Monday, October 16, 2017"
        );
        this.Test(
            new DateTime(2017, 10, 16, 0, 0, 0),
            "dddd,,, MMMM d,, yyyy,,,,",
            "Monday, October 16, 2017,"
        );
        this.Test(
            new DateTime(2020, 1, 1, 0, 35, 55),
            "m/d/yyyy\\ hh:mm:ss AM/PM;@",
            "1/1/2020 12:35:55 AM"
        );
        this.Test(
            new DateTime(2020, 1, 1, 12, 35, 55),
            "m/d/yyyy\\ hh:mm:ss AM/PM;@",
            "1/1/2020 12:35:55 PM"
        );
    }

    [Test]
    public void TestNumericDate1900()
    {
        this.Test("0", "dd/mm/yyyy", "0");
        this.Test(0, "dd/mm/yyyy", "00/01/1900");
        this.Test(0d, "dd/mm/yyyy", "00/01/1900");
        this.Test((short)0, "dd/mm/yyyy", "00/01/1900");
        this.Test(1, "dd/mm/yyyy", "01/01/1900");
        this.Test(60, "dd/mm/yyyy", "29/02/1900");
        this.Test(61, "dd/mm/yyyy", "01/03/1900");
        this.Test(43648, "[$-409]d\\-mmm\\-yyyy;@", "2-Jul-2019");
    }

    [Test]
    public void TestNumericDate1904()
    {
        this.Test(0, "dd/mm/yyyy", "01/01/1904", true);
        this.Test(0d, "dd/mm/yyyy", "01/01/1904", true);
        this.Test((short)0, "dd/mm/yyyy", "01/01/1904", true);
        this.Test(1, "dd/mm/yyyy", "02/01/1904", true);
        this.Test(60, "dd/mm/yyyy", "01/03/1904", true);
        this.Test(61, "dd/mm/yyyy", "02/03/1904", true);
    }

    [Test]
    public void TestNumericDuration()
    {
        this.Test(0, "[hh]:mm", "00:00");
        this.Test(1, "[hh]:mm", "24:00");
        this.Test(1.5, "[hh]:mm", "36:00");
    }

    [Test]
    public void TestTimeSpan()
    {
        this.Test(TimeSpan.FromHours(100), "[hh]:mm:ss", "100:00:00");
        this.Test(TimeSpan.FromHours(100), "[mm]:ss", "6000:00");
        this.Test(
            TimeSpan.FromMilliseconds(100 * 60 * 60 * 1000 + 123),
            "[mm]:ss.000",
            "6000:00.123"
        );

        this.Test(new TimeSpan(1, 2, 31, 45), "[hh]:mm:ss", "26:31:45");
        this.Test(new TimeSpan(1, 2, 31, 44, 500), "[hh]:mm:ss", "26:31:45");
        this.Test(new TimeSpan(1, 2, 31, 44, 500), "[hh]:mm:ss.000", "26:31:44.500");

        this.Test(new TimeSpan(-1, -2, -31, -45), "[hh]:mm:ss", "-26:31:45");
        this.Test(new TimeSpan(0, -2, -31, -45), "[hh]:mm:ss", "-02:31:45");
        this.Test(new TimeSpan(0, -2, -31, -44, -500), "[hh]:mm:ss", "-02:31:45");
        this.Test(new TimeSpan(0, -2, -31, -44, -500), "[hh]:mm:ss.000", "-02:31:44.500");
    }

    private void Test(object value, string format, string expected, bool isDate1904 = false)
    {
        string? result = Format(value, format, CultureInfo.InvariantCulture, isDate1904);
        ClassicAssert.AreEqual(expected, result);
    }

    [Test]
    public void TestFraction()
    {
        this.Test(1, "# ?/?", "1    ");
        this.Test(-1.2, "# ?/?", "-1 1/5");
        this.Test(12.3, "# ?/?", "12 1/3");
        this.Test(-12.34, "# ?/?", "-12 1/3");
        this.Test(123.45, "# ?/?", "123 4/9");
        this.Test(-123.456, "# ?/?", "-123 1/2");
        this.Test(1234.567, "# ?/?", "1234 4/7");
        this.Test(-1234.5678, "# ?/?", "-1234 4/7");
        this.Test(12345.6789, "# ?/?", "12345 2/3");
        this.Test(-12345.67891, "# ?/?", "-12345 2/3");

        this.Test(1, "# ??/??", "1      ");
        this.Test(-1.2, "# ??/??", "-1  1/5 ");
        this.Test(12.3, "# ??/??", "12  3/10");
        this.Test(-12.34, "# ??/??", "-12 17/50");
        this.Test(123.45, "# ??/??", "123  9/20");
        this.Test(-123.456, "# ??/??", "-123 26/57");
        this.Test(1234.567, "# ??/??", "1234 55/97");
        this.Test(-1234.5678, "# ??/??", "-1234 46/81");
        this.Test(12345.6789, "# ??/??", "12345 55/81");
        this.Test(-12345.67891, "# ??/??", "-12345 55/81");

        this.Test(1, "# ???/???", "1        ");
        this.Test(-1.2, "# ???/???", "-1   1/5  ");
        this.Test(12.3, "# ???/???", "12   3/10 ");
        this.Test(-12.34, "# ???/???", "-12  17/50 ");
        this.Test(123.45, "# ???/???", "123   9/20 ");
        this.Test(-123.456, "# ???/???", "-123  57/125");
        this.Test(1234.567, "# ???/???", "1234  55/97 ");
        this.Test(-1234.5678, "# ???/???", "-1234  67/118");
        this.Test(12345.6789, "# ???/???", "12345  74/109");
        this.Test(-12345.67891, "# ???/???", "-12345 573/844");

        this.Test(1, "# ?/2", "1    ");
        this.Test(-1.2, "# ?/2", "-1    ");
        this.Test(12.3, "# ?/2", "12 1/2");
        this.Test(-12.34, "# ?/2", "-12 1/2");
        this.Test(123.45, "# ?/2", "123 1/2");
        this.Test(-123.456, "# ?/2", "-123 1/2");
        this.Test(1234.567, "# ?/2", "1234 1/2");
        this.Test(-1234.5678, "# ?/2", "-1234 1/2");
        this.Test(12345.6789, "# ?/2", "12345 1/2");
        this.Test(-12345.67891, "# ?/2", "-12345 1/2");

        this.Test(1, "# ?/4", "1    ");
        this.Test(-1.2, "# ?/4", "-1 1/4");
        this.Test(12.3, "# ?/4", "12 1/4");
        this.Test(-12.34, "# ?/4", "-12 1/4");
        this.Test(123.45, "# ?/4", "123 2/4");
        this.Test(-123.456, "# ?/4", "-123 2/4");
        this.Test(1234.567, "# ?/4", "1234 2/4");
        this.Test(-1234.5678, "# ?/4", "-1234 2/4");
        this.Test(12345.6789, "# ?/4", "12345 3/4");
        this.Test(-12345.67891, "# ?/4", "-12345 3/4");

        this.Test(1, "# ?/8", "1    ");
        this.Test(-1.2, "# ?/8", "-1 2/8");
        this.Test(12.3, "# ?/8", "12 2/8");
        this.Test(-12.34, "# ?/8", "-12 3/8");
        this.Test(123.45, "# ?/8", "123 4/8");
        this.Test(-123.456, "# ?/8", "-123 4/8");
        this.Test(1234.567, "# ?/8", "1234 5/8");
        this.Test(-1234.5678, "# ?/8", "-1234 5/8");
        this.Test(12345.6789, "# ?/8", "12345 5/8");
        this.Test(-12345.67891, "# ?/8", "-12345 5/8");

        this.Test(1, "# ??/16", "1      ");
        this.Test(-1.2, "# ??/16", "-1  3/16");
        this.Test(12.3, "# ??/16", "12  5/16");
        this.Test(-12.34, "# ??/16", "-12  5/16");
        this.Test(123.45, "# ??/16", "123  7/16");
        this.Test(-123.456, "# ??/16", "-123  7/16");
        this.Test(1234.567, "# ??/16", "1234  9/16");
        this.Test(-1234.5678, "# ??/16", "-1234  9/16");
        this.Test(12345.6789, "# ??/16", "12345 11/16");
        this.Test(-12345.67891, "# ??/16", "-12345 11/16");

        this.Test(1, "# ?/10", "1     ");
        this.Test(-1.2, "# ?/10", "-1 2/10");
        this.Test(12.3, "# ?/10", "12 3/10");
        this.Test(-12.34, "# ?/10", "-12 3/10");
        this.Test(123.45, "# ?/10", "123 5/10");
        this.Test(-123.456, "# ?/10", "-123 5/10");
        this.Test(1234.567, "# ?/10", "1234 6/10");
        this.Test(-1234.5678, "# ?/10", "-1234 6/10");
        this.Test(12345.6789, "# ?/10", "12345 7/10");
        this.Test(-12345.67891, "# ?/10", "-12345 7/10");

        this.Test(1, "# ??/100", "1       ");
        this.Test(-1.2, "# ??/100", "-1 20/100");
        this.Test(12.3, "# ??/100", "12 30/100");
        this.Test(-12.34, "# ??/100", "-12 34/100");
        this.Test(123.45, "# ??/100", "123 45/100");
        this.Test(-123.456, "# ??/100", "-123 46/100");
        this.Test(1234.567, "# ??/100", "1234 57/100");
        this.Test(-1234.5678, "# ??/100", "-1234 57/100");
        this.Test(12345.6789, "# ??/100", "12345 68/100");
        this.Test(-12345.67891, "# ??/100", "-12345 68/100");

        this.Test(1, "??/??", " 1/1 ");
        this.Test(-1.2, "??/??", "- 6/5 ");
        this.Test(12.3, "??/??", "123/10");
        this.Test(-12.34, "??/??", "-617/50");
        this.Test(123.45, "??/??", "2469/20");
        this.Test(-123.456, "??/??", "-7037/57");
        this.Test(1234.567, "??/??", "119753/97");
        this.Test(-1234.5678, "??/??", "-100000/81");
        this.Test(12345.6789, "??/??", "1000000/81");
        this.Test(-12345.67891, "??/??", "-1000000/81");

        this.Test(0.3, "# ?/?", " 2/7");
        this.Test(1.3, "# ?/?", "1 1/3");
        this.Test(2.3, "# ?/?", "2 2/7");

        // Not sure what/why ssf does here:
        // Test(0.123251512342345, "# ??/?????????", "  480894/3901729");
        // Test(0.123251512342345, "# ?? / ?????????", "  480894 / 3901729");
        // This implementation instead renders like this:
        this.Test(0.123251512342345, "# ??/?????????", " 480894/3901729  ");
        this.Test(0.123251512342345, "# ?? / ?????????", " 480894 / 3901729  ");

        this.Test(0, "0", "0");
    }

    private void TestExponents(
        double value,
        string expected1,
        string expected2,
        string expected3,
        string expected4
    )
    {
        // value	#0.0E+0	##0.0E+0	###0.0E+0	####0.0E+0
        string? result1 = Format(value, "#0.0E+0", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected1, result1);

        string? result2 = Format(value, "##0.0E+0", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected2, result2);

        string? result3 = Format(value, "###0.0E+0", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected3, result3);

        string? result4 = Format(value, "####0.0E+0", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected4, result4);
    }

    private void TestNumber(
        double value,
        string expected1,
        string expected2,
        string expected3,
        string expected4,
        string expected5
    )
    {
        // value	?.?	??.??	???.???	???.?0?	???.?#?
        string? result1 = Format(value, "?.?", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected1, result1);

        string? result2 = Format(value, "??.??", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected2, result2);

        string? result3 = Format(value, "???.???", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected3, result3);

        string? result4 = Format(value, "???.?0?", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected4, result4);

        string? result5 = Format(value, "???.?#?", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected5, result5);
    }

    [Test]
    public void TestNumber()
    {
        this.TestNumber(0.0, " . ", "  .  ", "   .   ", "   . 0 ", "   .  ");
        this.TestNumber(0.1, " .1", "  .1 ", "   .1  ", "   .10 ", "   .1 ");
        this.TestNumber(0.12, " .1", "  .12", "   .12 ", "   .12 ", "   .12 ");
        this.TestNumber(0.123, " .1", "  .12", "   .123", "   .123", "   .123");

        this.TestNumber(1.0, "1. ", " 1.  ", "  1.   ", "  1. 0 ", "  1.  ");
        this.TestNumber(1.1, "1.1", " 1.1 ", "  1.1  ", "  1.10 ", "  1.1 ");
        this.TestNumber(1.12, "1.1", " 1.12", "  1.12 ", "  1.12 ", "  1.12 ");
        this.TestNumber(1.123, "1.1", " 1.12", "  1.123", "  1.123", "  1.123");
    }

    [Test]
    public void TestExponent()
    {
        this.TestExponents(-1.23457E-13, "-12.3E-14", "-123.5E-15", "-1234.6E-16", "-123.5E-15");
        this.TestExponents(-12345.6789, "-1.2E+4", "-12.3E+3", "-1.2E+4", "-12345.7E+0");

        this.TestExponents(1.23457E-13, "12.3E-14", "123.5E-15", "1234.6E-16", "123.5E-15");
        this.TestExponents(1.23457E-12, "1.2E-12", "1.2E-12", "1.2E-12", "1234.6E-15");
        this.TestExponents(1.23457E-11, "12.3E-12", "12.3E-12", "12.3E-12", "12345.7E-15");
        this.TestExponents(1.23457E-10, "1.2E-10", "123.5E-12", "123.5E-12", "1.2E-10");
        this.TestExponents(1.23457E-09, "12.3E-10", "1.2E-9", "1234.6E-12", "12.3E-10");
        this.TestExponents(1.23457E-08, "1.2E-8", "12.3E-9", "1.2E-8", "123.5E-10");
        this.TestExponents(0.000000123457, "12.3E-8", "123.5E-9", "12.3E-8", "1234.6E-10");
        this.TestExponents(0.00000123457, "1.2E-6", "1.2E-6", "123.5E-8", "12345.7E-10");
        this.TestExponents(0.0000123457, "12.3E-6", "12.3E-6", "1234.6E-8", "1.2E-5");
        this.TestExponents(0.000123457, "1.2E-4", "123.5E-6", "1.2E-4", "12.3E-5");
        this.TestExponents(0.001234568, "12.3E-4", "1.2E-3", "12.3E-4", "123.5E-5");
        this.TestExponents(0.012345679, "1.2E-2", "12.3E-3", "123.5E-4", "1234.6E-5");
        this.TestExponents(0.123456789, "12.3E-2", "123.5E-3", "1234.6E-4", "12345.7E-5");
        this.TestExponents(1.23456789, "1.2E+0", "1.2E+0", "1.2E+0", "1.2E+0");
        this.TestExponents(12.3456789, "12.3E+0", "12.3E+0", "12.3E+0", "12.3E+0");
        this.TestExponents(123.456789, "1.2E+2", "123.5E+0", "123.5E+0", "123.5E+0");
        this.TestExponents(1234.56789, "12.3E+2", "1.2E+3", "1234.6E+0", "1234.6E+0");
        this.TestExponents(12345.6789, "1.2E+4", "12.3E+3", "1.2E+4", "12345.7E+0");
        this.TestExponents(123456.789, "12.3E+4", "123.5E+3", "12.3E+4", "1.2E+5");
        this.TestExponents(1234567.89, "1.2E+6", "1.2E+6", "123.5E+4", "12.3E+5");
        this.TestExponents(12345678.9, "12.3E+6", "12.3E+6", "1234.6E+4", "123.5E+5");
        this.TestExponents(123456789D, "1.2E+8", "123.5E+6", "1.2E+8", "1234.6E+5");
        this.TestExponents(1234567890D, "12.3E+8", "1.2E+9", "12.3E+8", "12345.7E+5");
        this.TestExponents(12345678900D, "1.2E+10", "12.3E+9", "123.5E+8", "1.2E+10");
        this.TestExponents(123456789000D, "12.3E+10", "123.5E+9", "1234.6E+8", "12.3E+10");
        this.TestExponents(1234567890000D, "1.2E+12", "1.2E+12", "1.2E+12", "123.5E+10");
        this.TestExponents(12345678900000D, "12.3E+12", "12.3E+12", "12.3E+12", "1234.6E+10");
        this.TestExponents(123456789000000D, "1.2E+14", "123.5E+12", "123.5E+12", "12345.7E+10");
        this.TestExponents(1234567890000000D, "12.3E+14", "1.2E+15", "1234.6E+12", "1.2E+15");
        this.TestExponents(12345678900000000D, "1.2E+16", "12.3E+15", "1.2E+16", "12.3E+15");
        this.TestExponents(123456789000000000D, "12.3E+16", "123.5E+15", "12.3E+16", "123.5E+15");
        this.TestExponents(1234567890000000000D, "1.2E+18", "1.2E+18", "123.5E+16", "1234.6E+15");
        this.TestExponents(
            12345678900000000000D,
            "12.3E+18",
            "12.3E+18",
            "1234.6E+16",
            "12345.7E+15"
        );
        this.TestExponents(123456789000000000000D, "1.2E+20", "123.5E+18", "1.2E+20", "1.2E+20");
        this.TestExponents(1234567890000000000000D, "12.3E+20", "1.2E+21", "12.3E+20", "12.3E+20");
        this.TestExponents(
            12345678900000000000000D,
            "1.2E+22",
            "12.3E+21",
            "123.5E+20",
            "123.5E+20"
        );
        this.TestExponents(
            123456789000000000000000D,
            "12.3E+22",
            "123.5E+21",
            "1234.6E+20",
            "1234.6E+20"
        );
        this.TestExponents(
            1234567890000000000000000D,
            "1.2E+24",
            "1.2E+24",
            "1.2E+24",
            "12345.7E+20"
        );
        this.TestExponents(
            12345678900000000000000000D,
            "12.3E+24",
            "12.3E+24",
            "12.3E+24",
            "1.2E+25"
        );
        this.TestExponents(
            123456789000000000000000000D,
            "1.2E+26",
            "123.5E+24",
            "123.5E+24",
            "12.3E+25"
        );
        this.TestExponents(
            1234567890000000000000000000D,
            "12.3E+26",
            "1.2E+27",
            "1234.6E+24",
            "123.5E+25"
        );
        this.TestExponents(
            12345678900000000000000000000D,
            "1.2E+28",
            "12.3E+27",
            "1.2E+28",
            "1234.6E+25"
        );
        this.TestExponents(
            123456789000000000000000000000D,
            "12.3E+28",
            "123.5E+27",
            "12.3E+28",
            "12345.7E+25"
        );
        this.TestExponents(
            1234567890000000000000000000000D,
            "1.2E+30",
            "1.2E+30",
            "123.5E+28",
            "1.2E+30"
        );
        this.TestExponents(
            12345678900000000000000000000000D,
            "12.3E+30",
            "12.3E+30",
            "1234.6E+28",
            "12.3E+30"
        );
    }

    private void TestComma(
        double value,
        string expected1,
        string expected2,
        string expected3,
        string expected4,
        string expected5,
        string expected6,
        string expected7
    )
    {
        // value	#.0000,,,	#.0000,,	#.0000,	#,##0.0	###,##0	###,###	#,###.00
        string? result1 = Format(value, "#.0000,,,", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected1, result1);

        string? result2 = Format(value, "#.0000,,", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected2, result2);

        string? result3 = Format(value, "#.0000,", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected3, result3);

        string? result4 = Format(value, "#,##0.0", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected4, result4);

        string? result5 = Format(value, "###,##0", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected5, result5);

        string? result6 = Format(value, "###,###", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected6, result6);

        string? result7 = Format(value, "#,###.00", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual(expected7, result7);
    }

    [Test]
    public void TestComma()
    {
        this.TestComma(0.99, ".0000", ".0000", ".0010", "1.0", "1", "1", ".99");
        this.TestComma(1.2345, ".0000", ".0000", ".0012", "1.2", "1", "1", "1.23");
        this.TestComma(12.345, ".0000", ".0000", ".0123", "12.3", "12", "12", "12.35");
        this.TestComma(123.456, ".0000", ".0001", ".1235", "123.5", "123", "123", "123.46");
        this.TestComma(1234, ".0000", ".0012", "1.2340", "1,234.0", "1,234", "1,234", "1,234.00");
        this.TestComma(
            12345,
            ".0000",
            ".0123",
            "12.3450",
            "12,345.0",
            "12,345",
            "12,345",
            "12,345.00"
        );
        this.TestComma(
            123456,
            ".0001",
            ".1235",
            "123.4560",
            "123,456.0",
            "123,456",
            "123,456",
            "123,456.00"
        );
        this.TestComma(
            1234567,
            ".0012",
            "1.2346",
            "1234.5670",
            "1,234,567.0",
            "1,234,567",
            "1,234,567",
            "1,234,567.00"
        );
        this.TestComma(
            12345678,
            ".0123",
            "12.3457",
            "12345.6780",
            "12,345,678.0",
            "12,345,678",
            "12,345,678",
            "12,345,678.00"
        );
        this.TestComma(
            123456789,
            ".1235",
            "123.4568",
            "123456.7890",
            "123,456,789.0",
            "123,456,789",
            "123,456,789",
            "123,456,789.00"
        );
        this.TestComma(
            1234567890,
            "1.2346",
            "1234.5679",
            "1234567.8900",
            "1,234,567,890.0",
            "1,234,567,890",
            "1,234,567,890",
            "1,234,567,890.00"
        );
        this.TestComma(
            12345678901,
            "12.3457",
            "12345.6789",
            "12345678.9010",
            "12,345,678,901.0",
            "12,345,678,901",
            "12,345,678,901",
            "12,345,678,901.00"
        );
        this.TestComma(
            123456789012,
            "123.4568",
            "123456.7890",
            "123456789.0120",
            "123,456,789,012.0",
            "123,456,789,012",
            "123,456,789,012",
            "123,456,789,012.00"
        );
        this.TestComma(4321, ".0000", ".0043", "4.3210", "4,321.0", "4,321", "4,321", "4,321.00");
        this.TestComma(
            4321234,
            ".0043",
            "4.3212",
            "4321.2340",
            "4,321,234.0",
            "4,321,234",
            "4,321,234",
            "4,321,234.00"
        );
    }

    [Test]
    public void TestThousandSeparator()
    {
        string? actual = Format(1469.07, "0,000,000.00", CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual("0,001,469.07", actual);
    }

    [Test]
    public void TestThousandSeparatorCulture()
    {
        string? actual = Format(1469.07, "0,000,000.00", new CultureInfo("da-DK"));
        ClassicAssert.AreEqual("0.001.469,07", actual);
    }

    [Test]
    [Arguments("da-DK", "17.08.1978")]
    [Arguments("en-US", "17/08/1978")]
    [Arguments("bg-BG", "17.08.1978")]
    [Arguments("nb-NO", "17.08.1978")]
    public void TestDateSeparatorCulture(string cultureName, string expected)
    {
        string? actual = Format(
            new DateTime(1978, 8, 17),
            "DD/MM/YYYY",
            new CultureInfo(cultureName)
        );
        ClassicAssert.AreEqual(expected, actual);
    }

    private void TestValid(string format)
    {
        NumberFormat to = new(format);
        ClassicAssert.IsTrue(to.IsValid, $"Invalid format: {format}");
    }

    private void TestInvalid(string format)
    {
        NumberFormat to = new(format);
        ClassicAssert.IsFalse(to.IsValid, $"Expected invalid format: {format}");
    }

    [Test]
    public void TestInvalid()
    {
        this.TestInvalid("\"");
        this.TestInvalid("[abc");
        this.TestInvalid("~");
        this.TestInvalid("#~0");
    }

    [Test]
    public void TestValid()
    {
        this.TestValid("\" Excellent\"");
        this.TestValid("\" Fair\"");
        this.TestValid("\" Good\"");
        this.TestValid("\" Poor\"");
        this.TestValid("\" Very Good\"");
        this.TestValid("\"$\"#,##0");
        this.TestValid("\"$\"#,##0.00");
        this.TestValid("\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)");
        this.TestValid("\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)");
        this.TestValid("\"$\"#,##0;[Red]\\-\"$\"#,##0");
        this.TestValid("\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)");
        this.TestValid("\"$\"#,##0_);\\(\"$\"#,##0\\)");
        this.TestValid("\"Haha!\"\\ @\\ \"Yeah!\"");
        this.TestValid("\"TRUE\";\"TRUE\";\"FALSE\"");
        this.TestValid("\"True\";\"True\";\"False\";@");
        this.TestValid("\"Years: \"0");
        this.TestValid("\"Yes\";\"Yes\";\"No\";@");
        this.TestValid("\"kl \"hh:mm:ss;@");
        this.TestValid("\"£\"#,##0.00");
        this.TestValid("\"£\"#,##0;[Red]\\-\"£\"#,##0");
        this.TestValid("\"€\"#,##0.00");
        this.TestValid("\"€\"\\ #,##0.00_-");
        this.TestValid("\"上午/下午 \"hh\"時\"mm\"分\"ss\"秒 \"");
        this.TestValid("\"￥\"#,##0.00;\"￥\"\\-#,##0.00");
        this.TestValid("#");
        this.TestValid("# ?/?");
        this.TestValid("# ??/??");
        this.TestValid("#\" \"?/?");
        this.TestValid("#\" \"??/??");
        this.TestValid("#\"abded\"\\ ??/??");
        this.TestValid("###0.00;-###0.00");
        this.TestValid("###0;-###0");
        this.TestValid("##0.0E+0");
        this.TestValid("#,##0");
        this.TestValid("#,##0 ;(#,##0)");
        this.TestValid("#,##0 ;[Red](#,##0)");
        this.TestValid("#,##0\"р.\";[Red]\\-#,##0\"р.\"");
        this.TestValid("#,##0.0");
        this.TestValid("#,##0.00");
        this.TestValid("#,##0.00 \"�\"");
        this.TestValid("#,##0.00 €;-#,##0.00 €");
        this.TestValid("#,##0.00\"р.\";[Red]\\-#,##0.00\"р.\"");
        this.TestValid("#,##0.000");
        this.TestValid("#,##0.0000");
        this.TestValid("#,##0.00000");
        this.TestValid("#,##0.000000");
        this.TestValid("#,##0.0000000");
        this.TestValid("#,##0.00000000");
        this.TestValid("#,##0.000000000");
        this.TestValid("#,##0.00000000;[Red]#,##0.00000000");
        this.TestValid("#,##0.0000_ ");
        this.TestValid("#,##0.000_ ");
        this.TestValid("#,##0.000_);\\(#,##0.000\\)");
        this.TestValid("#,##0.00;(#,##0.00)");
        this.TestValid("#,##0.00;(#,##0.00);0.00");
        this.TestValid("#,##0.00;[Red](#,##0.00)");
        this.TestValid("#,##0.00;[Red]\\(#,##0.00\\)");
        this.TestValid("#,##0.00;\\(#,##0.00\\)");
        this.TestValid("#,##0.00[$₹-449]_);\\(#,##0.00[$₹-449]\\)");
        this.TestValid("#,##0.00\\ \"р.\"");
        this.TestValid("#,##0.00\\ \"р.\";[Red]\\-#,##0.00\\ \"р.\"");
        this.TestValid("#,##0.00\\ [$€-407]");
        this.TestValid("#,##0.00\\ [$€-40C]");
        this.TestValid("#,##0.00_);\\(#,##0.00\\)");
        this.TestValid("#,##0.00_р_.;[Red]\\-#,##0.00_р_.");
        this.TestValid("#,##0.00_р_.;\\-#,##0.00_р_.");
        this.TestValid("#,##0.0;[Red]#,##0.0");
        this.TestValid("#,##0.0_ ;\\-#,##0.0\\ ");
        this.TestValid("#,##0.0_);[Red]\\(#,##0.0\\)");
        this.TestValid("#,##0.0_);\\(#,##0.0\\)");
        this.TestValid("#,##0;\\-#,##0;0");
        this.TestValid("#,##0\\ \"р.\";[Red]\\-#,##0\\ \"р.\"");
        this.TestValid("#,##0\\ \"р.\";\\-#,##0\\ \"р.\"");
        this.TestValid("#,##0\\ ;[Red]\\(#,##0\\)");
        this.TestValid("#,##0\\ ;\\(#,##0\\)");
        this.TestValid("#,##0_ ");
        this.TestValid("#,##0_ ;[Red]\\-#,##0\\ ");
        this.TestValid("#,##0_);[Red]\\(#,##0\\)");
        this.TestValid("#,##0_р_.;[Red]\\-#,##0_р_.");
        this.TestValid("#,##0_р_.;\\-#,##0_р_.");
        this.TestValid("#.0000,,");
        this.TestValid("#0");
        this.TestValid("#0.00");
        this.TestValid("#0.0000");
        this.TestValid("#\\ ?/10");
        this.TestValid("#\\ ?/2");
        this.TestValid("#\\ ?/4");
        this.TestValid("#\\ ?/8");
        this.TestValid("#\\ ?/?");
        this.TestValid("#\\ ??/100");
        this.TestValid("#\\ ??/100;[Red]\\(#\\ ??/16\\)");
        this.TestValid("#\\ ??/16");
        this.TestValid("#\\ ??/??");
        this.TestValid("#\\ ??/?????????");
        this.TestValid("#\\ ???/???");
        this.TestValid("**\\ #,###,#00,000.00,**");
        this.TestValid("0");
        this.TestValid("0\"abde\".0\"??\"000E+00");
        this.TestValid("0%");
        this.TestValid("0.0");
        this.TestValid("0.0%");
        this.TestValid("0.00");
        this.TestValid("0.00\"°\"");
        this.TestValid("0.00%");
        this.TestValid("0.000");
        this.TestValid("0.000%");
        this.TestValid("0.0000");
        this.TestValid("0.000000");
        this.TestValid("0.00000000");
        this.TestValid("0.000000000");
        this.TestValid("0.000000000%");
        this.TestValid("0.00000000000");
        this.TestValid("0.000000000000000");
        this.TestValid("0.00000000E+00");
        this.TestValid("0.0000E+00");
        this.TestValid("0.00;[Red]0.00");
        this.TestValid("0.00E+00");
        this.TestValid("0.00_);[Red]\\(0.00\\)");
        this.TestValid("0.00_);\\(0.00\\)");
        this.TestValid("0.0_ ");
        this.TestValid("00.00.00.000");
        this.TestValid("00.000%");
        this.TestValid("0000");
        this.TestValid("00000");
        this.TestValid("00000000");
        this.TestValid("000000000");
        this.TestValid("00000\\-0000");
        this.TestValid("00000\\-00000");
        this.TestValid("000\\-00\\-0000");
        this.TestValid("0;[Red]0");
        this.TestValid("0\\-00000\\-00000\\-0");
        this.TestValid("0_);[Red]\\(0\\)");
        this.TestValid("0_);\\(0\\)");
        this.TestValid("@");
        this.TestValid("A/P");
        this.TestValid("AM/PM");
        this.TestValid("AM/PMh\"時\"mm\"分\"ss\"秒\";@");
        this.TestValid("D");
        this.TestValid("DD");
        this.TestValid("DD/MM/YY;@");
        this.TestValid("DD/MM/YYYY");
        this.TestValid("DD/MM/YYYY;@");
        this.TestValid("DDD");
        this.TestValid("DDDD");
        this.TestValid("DDDD\", \"MMMM\\ DD\", \"YYYY");
        this.TestValid("GENERAL");
        this.TestValid("General");
        this.TestValid("H");
        this.TestValid("H:MM:SS\\ AM/PM");
        this.TestValid("HH:MM");
        this.TestValid("HH:MM:SS\\ AM/PM");
        this.TestValid("HHM");
        this.TestValid("HHMM");
        this.TestValid("HH[MM]");
        this.TestValid("HH[M]");
        this.TestValid("M/D/YYYY");
        this.TestValid("M/D/YYYY\\ H:MM");
        this.TestValid("MM/DD/YY");
        this.TestValid("S");
        this.TestValid("SS");
        this.TestValid("YY");
        this.TestValid("YYM");
        this.TestValid("YYMM");
        this.TestValid("YYMMM");
        this.TestValid("YYMMMM");
        this.TestValid("YYMMMMM");
        this.TestValid("YYYY");
        this.TestValid("YYYY-MM-DD HH:MM:SS");
        this.TestValid("YYYY\\-MM\\-DD");
        this.TestValid("[$$-409]#,##0");
        this.TestValid("[$$-409]#,##0.00");
        this.TestValid("[$$-409]#,##0.00_);[Red]\\([$$-409]#,##0.00\\)");
        this.TestValid("[$$-C09]#,##0.00");
        this.TestValid("[$-100042A]h:mm:ss\\ AM/PM;@");
        this.TestValid("[$-1010409]0.000%");
        this.TestValid("[$-1010409]General");
        this.TestValid("[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@");
        this.TestValid("[$-1010409]dddd, mmmm dd, yyyy");
        this.TestValid("[$-1010409]m/d/yyyy");
        this.TestValid("[$-1409]h:mm:ss\\ AM/PM;@");
        this.TestValid("[$-2000000]h:mm:ss;@");
        this.TestValid("[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@");
        this.TestValid("[$-4000439]h:mm:ss\\ AM/PM;@");
        this.TestValid("[$-4010439]d/m/yyyy\\ h:mm\\ AM/PM;@");
        this.TestValid("[$-409]AM/PM\\ hh:mm:ss;@");
        this.TestValid("[$-409]d/m/yyyy\\ hh:mm;@");
        this.TestValid("[$-409]d\\-mmm;@");
        this.TestValid("[$-409]d\\-mmm\\-yy;@");
        this.TestValid("[$-409]d\\-mmm\\-yyyy;@");
        this.TestValid("[$-409]dd/mm/yyyy\\ hh:mm;@");
        this.TestValid("[$-409]dd\\-mmm\\-yy;@");
        this.TestValid("[$-409]h:mm:ss\\ AM/PM;@");
        this.TestValid("[$-409]h:mm\\ AM/PM;@");
        this.TestValid("[$-409]m/d/yy\\ h:mm\\ AM/PM;@");
        this.TestValid("[$-409]mmm\\-yy;@");
        this.TestValid("[$-409]mmmm\\ d\\,\\ yyyy;@");
        this.TestValid("[$-409]mmmm\\-yy;@");
        this.TestValid("[$-409]mmmmm;@");
        this.TestValid("[$-409]mmmmm\\-yy;@");
        this.TestValid("[$-40E]h\\ \"óra\"\\ m\\ \"perckor\"\\ AM/PM;@");
        this.TestValid("[$-412]AM/PM\\ h\"시\"\\ mm\"분\"\\ ss\"초\";@");
        this.TestValid("[$-41C]h:mm:ss\\.AM/PM;@");
        this.TestValid("[$-449]hh:mm:ss\\ AM/PM;@");
        this.TestValid("[$-44E]hh:mm:ss\\ AM/PM;@");
        this.TestValid("[$-44F]hh:mm:ss\\ AM/PM;@");
        this.TestValid("[$-D000409]h:mm\\ AM/PM;@");
        this.TestValid("[$-D010000]d/mm/yyyy\\ h:mm\\ \"น.\";@");
        this.TestValid("[$-F400]h:mm:ss\\ AM/PM");
        this.TestValid("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");
        this.TestValid("[$AUD]\\ #,##0.00");
        this.TestValid("[$RD$-1C0A]#,##0.00;[Red]\\-[$RD$-1C0A]#,##0.00");
        this.TestValid("[$SFr.-810]\\ #,##0.00_);[Red]\\([$SFr.-810]\\ #,##0.00\\)");
        this.TestValid("[$£-809]#,##0.00;[Red][$£-809]#,##0.00");
        this.TestValid("[$¥-411]#,##0.00");
        this.TestValid("[$¥-804]#,##0.00");
        this.TestValid("[<0]\"\";0%");
        this.TestValid("[<=9999999]###\\-####;\\(###\\)\\ ###\\-####");
        this.TestValid("[=0]?;#,##0.00");
        this.TestValid("[=0]?;0%");
        this.TestValid("[=0]?;[<4.16666666666667][hh]:mm:ss;[hh]:mm");
        this.TestValid("[>999999]#,,\"M\";[>999]#,\"K\";#");
        this.TestValid("[>999999]#.000,,\"M\";[>999]#.000,\"K\";#.000");
        this.TestValid("[>=100000]0.000\\ \\\";[Red]0.000\\ \\<\\ \\>\\ \\\"\\ \\&\\ \\'\\ ");
        this.TestValid("[>=100000]0.000\\ \\<;[Red]0.000\\ \\>");
        this.TestValid("[BLACK]@");
        this.TestValid("[BLUE]GENERAL");
        this.TestValid("[Black]@");
        this.TestValid("[Blue]General");
        this.TestValid("[CYAN]@");
        this.TestValid("[Cyan]@");
        this.TestValid("[DBNum1][$-804]AM/PMh\"时\"mm\"分\";@");
        this.TestValid("[DBNum1][$-804]General");
        this.TestValid("[DBNum1][$-804]h\"时\"mm\"分\";@");
        this.TestValid("[ENG][$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@");
        this.TestValid("[ENG][$-101040D]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-101042A]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-140C]dddd\\ \"YeahWoo!\"\\ ddd\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-2C0A]dddd\\ d\" de \"mmmm\" de \"yyyy;@");
        this.TestValid("[ENG][$-402]dd\\ mmmm\\ yyyy\\ \"г.\";@");
        this.TestValid("[ENG][$-403]dddd\\,\\ d\" / \"mmmm\" / \"yyyy;@");
        this.TestValid("[ENG][$-405]d\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-408]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-409]d\\-mmm;@");
        this.TestValid("[ENG][$-409]d\\-mmm\\-yy;@");
        this.TestValid("[ENG][$-409]d\\-mmm\\-yyyy;@");
        this.TestValid("[ENG][$-409]dd\\-mmm\\-yy;@");
        this.TestValid("[ENG][$-409]mmm\\-yy;@");
        this.TestValid("[ENG][$-409]mmmm\\ d\\,\\ yyyy;@");
        this.TestValid("[ENG][$-409]mmmm\\-yy;@");
        this.TestValid("[ENG][$-40B]d\\.\\ mmmm\\t\\a\\ yyyy;@");
        this.TestValid("[ENG][$-40C]d/mmm/yyyy;@");
        this.TestValid("[ENG][$-40E]yyyy/\\ mmmm\\ d\\.;@");
        this.TestValid("[ENG][$-40F]dd\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-410]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-415]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-416]d\\ \\ mmmm\\,\\ yyyy;@");
        this.TestValid("[ENG][$-418]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-41A]d\\.\\ mmmm\\ yyyy\\.;@");
        this.TestValid("[ENG][$-41B]d\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-41D]\"den \"\\ d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-420]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@");
        this.TestValid("[ENG][$-421]dd\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-424]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-425]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-426]dddd\\,\\ yyyy\". gada \"d\\.\\ mmmm;@");
        this.TestValid("[ENG][$-427]yyyy\\ \"m.\"\\ mmmm\\ d\\ \"d.\";@");
        this.TestValid("[ENG][$-42B]dddd\\,\\ d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-42C]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-42D]yyyy\"(e)ko\"\\ mmmm\"ren\"\\ d\"a\";@");
        this.TestValid("[ENG][$-42F]dddd\\,\\ dd\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-437]yyyy\\ \\წ\\ლ\\ი\\ს\\ dd\\ mm\\,\\ dddd;@");
        this.TestValid("[ENG][$-438]d\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-43F]d\\ mmmm\\ yyyy\\ \"ж.\";@");
        this.TestValid("[ENG][$-444]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-449]dd\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-44E]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-44F]dd\\ mmmm\\ yyyy\\ dddd;@");
        this.TestValid("[ENG][$-457]dd\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-813]dddd\\ d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-81A]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-82C]d\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-843]yyyy\\ \"й\"\"и\"\"л\"\\ d/mmmm;@");
        this.TestValid("[ENG][$-C07]dddd\\,\\ dd\\.\\ mmmm\\ yyyy;@");
        this.TestValid("[ENG][$-FC19]yyyy\\,\\ dd\\ mmmm;@");
        this.TestValid("[ENG][$-FC22]d\\ mmmm\\ yyyy\" р.\";@");
        this.TestValid("[ENG][$-FC23]d\\ mmmm\\ yyyy;@");
        this.TestValid("[GREEN]#,###");
        this.TestValid("[Green]#,###");
        this.TestValid("[HH]");
        this.TestValid("[HIJ][$-2060401]d/mm/yyyy\\ h:mm\\ AM/PM;@");
        this.TestValid("[HIJ][$-2060401]d\\ mmmm\\ yyyy;@");
        this.TestValid("[H]");
        this.TestValid("[JPN][$-411]gggyy\"年\"m\"月\"d\"日\"\\ dddd;@");
        this.TestValid("[MAGENTA]0.00");
        this.TestValid("[Magenta]0.00");
        this.TestValid("[RED]#.##");
        this.TestValid("[Red]#.##");
        this.TestValid("[Red][<-25]General;[Blue][>25]General;[Green]General;[Yellow]General\\ ");
        this.TestValid("[Red][<=-25]General;[Blue][>=25]General;[Green]General;[Yellow]General");
        this.TestValid("[Red][<>50]General;[Blue]000");
        this.TestValid("[Red][=50]General;[Blue]000");
        this.TestValid("[SS]");
        this.TestValid("[S]");
        this.TestValid("[TWN][DBNum1][$-404]y\"年\"m\"月\"d\"日\";@");
        this.TestValid("[WHITE]0.0");
        this.TestValid("[White]0.0");
        this.TestValid("[YELLOW]@");
        this.TestValid("[Yellow]@");
        this.TestValid("[h]");
        this.TestValid("[h]:mm:ss");
        this.TestValid("[h]:mm:ss;@");
        this.TestValid("[h]\\.mm\" Uhr \";@");
        this.TestValid("[hh]");
        this.TestValid("[s]");
        this.TestValid("[ss]");
        this.TestValid("\\#\\r\\e\\c");
        this.TestValid("\\$#,##0_);[Red]\"($\"#,##0\\)");
        this.TestValid("\\$0.00");
        this.TestValid("\\C\\O\\B\\ \\o\\n\\ @");
        this.TestValid("\\C\\R\\O\\N\\T\\A\\B\\ \\o\\n\\ @");
        this.TestValid("\\R\\e\\s\\u\\l\\t\\ \\o\\n\\ @");
        this.TestValid("\\S\\Q\\L\\ \\:\\ @");
        this.TestValid("\\S\\Q\\L\\ \\R\\e\\q\\u\\e\\s\\t\\ \\f\\o\\r\\ @");
        this.TestValid("\\c\\c\\c?????0\"aaaa\"0\"bbbb\"000000.00%");
        this.TestValid("\\u\\n\\t\\i\\l\\ h:mm;@");
        this.TestValid(
            "_ \"￥\"* #,##0.00_ \"Positive\";_ \"￥\"* \\-#,##0.00_ ;_ \"￥\"* \"-\"??_ \"Negtive\";_ @_ \\ \"Zero\""
        );
        this.TestValid(
            "_ * #,##0.00_)[$﷼-429]_ ;_ * \\(#,##0.00\\)[$﷼-429]_ ;_ * \"-\"??_)[$﷼-429]_ ;_ @_ "
        );
        this.TestValid("_ * #,##0_ ;_ * \\-#,##0_ ;[Red]_ * \"-\"_ ;_ @_ ");
        this.TestValid("_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)");
        this.TestValid("_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"??_);_(@_)");
        this.TestValid("_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)");
        this.TestValid("_(* #,##0.0000_);_(* \\(#,##0.0000\\);_(* \"-\"??_);_(@_)");
        this.TestValid("_(* #,##0.000_);_(* \\(#,##0.000\\);_(* \"-\"??_);_(@_)");
        this.TestValid("_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)");
        this.TestValid("_(* #,##0.0_);_(* \\(#,##0.0\\);_(* \"-\"??_);_(@_)");
        this.TestValid("_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"??_);_(@_)");
        this.TestValid("_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)");
        this.TestValid(
            "_([$ANG]\\ * #,##0.0_);_([$ANG]\\ * \\(#,##0.0\\);_([$ANG]\\ * \"-\"?_);_(@_)"
        );
        this.TestValid(
            "_-\"€\"\\ * #,##0.00_-;_-\"€\"\\ * #,##0.00\\-;_-\"€\"\\ * \"-\"??_-;_-@_-"
        );
        this.TestValid("_-* #,##0.00\" TL\"_-;\\-* #,##0.00\" TL\"_-;_-* \\-??\" TL\"_-;_-@_-");
        this.TestValid("_-* #,##0.00\" €\"_-;\\-* #,##0.00\" €\"_-;_-* \\-??\" €\"_-;_-@_-");
        this.TestValid(
            "_-* #,##0.00\\ \"р.\"_-;\\-* #,##0.00\\ \"р.\"_-;_-* \"-\"??\\ \"р.\"_-;_-@_-"
        );
        this.TestValid(
            "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-"
        );
        this.TestValid(
            "_-* #,##0.00\\ [$€-407]_-;\\-* #,##0.00\\ [$€-407]_-;_-* \\-??\\ [$€-407]_-;_-@_-"
        );
        this.TestValid("_-* #,##0.0\\ _F_-;\\-* #,##0.0\\ _F_-;_-* \"-\"??\\ _F_-;_-@_-");
        this.TestValid("_-* #,##0\\ \"€\"_-;\\-* #,##0\\ \"€\"_-;_-* \"-\"\\ \"€\"_-;_-@_-");
        this.TestValid("_-* #,##0_-;\\-* #,##0_-;_-* \"-\"??_-;_-@_-");
        this.TestValid("_-\\$* #,##0.0_ ;_-\\$* \\-#,##0.0\\ ;_-\\$* \"-\"?_ ;_-@_ ");
        this.TestValid("d");
        this.TestValid("d-mmm");
        this.TestValid("d-mmm-yy");
        this.TestValid("d/m");
        this.TestValid("d/m/yy;@");
        this.TestValid("d/m/yyyy;@");
        this.TestValid("d/mm/yy;@");
        this.TestValid("d/mm/yyyy;@");
        this.TestValid("d\\-mmm");
        this.TestValid("d\\-mmm\\-yyyy");
        this.TestValid("dd");
        this.TestValid("dd\"-\"mmm\"-\"yyyy");
        this.TestValid("dd/m/yyyy");
        this.TestValid("dd/mm/yy");
        this.TestValid("dd/mm/yy;@");
        this.TestValid("dd/mm/yy\\ hh:mm");
        this.TestValid("dd/mm/yyyy");
        this.TestValid("dd/mm/yyyy\\ hh:mm:ss");
        this.TestValid("dd/mmm");
        this.TestValid("dd\\-mm\\-yy");
        this.TestValid("dd\\-mmm\\-yy");
        this.TestValid("dd\\-mmm\\-yyyy\\ hh:mm:ss.000");
        this.TestValid("dd\\/mm\\/yy");
        this.TestValid("dd\\/mm\\/yyyy");
        this.TestValid("ddd");
        this.TestValid("dddd");
        this.TestValid("dddd, mmmm dd, yyyy");
        this.TestValid("h");
        this.TestValid("h\"时\"mm\"分\"ss\"秒\";@");
        this.TestValid("h\"時\"mm\"分\"ss\"秒\";@");
        this.TestValid("h:mm");
        this.TestValid("h:mm AM/PM");
        this.TestValid("h:mm:ss");
        this.TestValid("h:mm:ss AM/PM");
        this.TestValid("h:mm:ss;@");
        this.TestValid("h:mm;@");
        this.TestValid("h\\.mm\" Uhr \";@");
        this.TestValid("h\\.mm\" h\";@");
        this.TestValid("h\\.mm\" u.\";@");
        this.TestValid("hh\":\"mm AM/PM");
        this.TestValid("hh:mm:ss");
        this.TestValid("hh:mm:ss\\ AM/PM");
        this.TestValid("hh\\.mm\" h\";@");
        this.TestValid("hhm");
        this.TestValid("hhmm");
        this.TestValid("m\"月\"d\"日\"");
        this.TestValid("m/d/yy");
        this.TestValid("m/d/yy h:mm");
        this.TestValid("m/d/yy;@");
        this.TestValid("m/d/yy\\ h:mm");
        this.TestValid("m/d/yy\\ h:mm;@");
        this.TestValid("m/d/yyyy");
        this.TestValid("m/d/yyyy;@");
        this.TestValid("m/d/yyyy\\ h:mm:ss;@");
        this.TestValid("m/d;@");
        this.TestValid("m\\/d\\/yyyy");
        this.TestValid("mm/dd");
        this.TestValid("mm/dd/yy");
        this.TestValid("mm/dd/yy;@");
        this.TestValid("mm/dd/yyyy");
        this.TestValid("mm:ss");
        this.TestValid("mm:ss.0;@");
        this.TestValid("mmm d, yyyy");
        this.TestValid("mmm\" \"d\", \"yyyy");
        this.TestValid("mmm-yy");
        this.TestValid("mmm-yy;@");
        this.TestValid("mmm/yy");
        this.TestValid("mmm\\-yy");
        this.TestValid("mmm\\-yy;@");
        this.TestValid("mmm\\-yyyy");
        this.TestValid("mmmm\\ d\\,\\ yyyy");
        this.TestValid("mmmm\\ yyyy");
        this.TestValid("mmss.0");
        this.TestValid("s");
        this.TestValid("ss");
        this.TestValid("yy");
        this.TestValid("yy/mm/dd");
        this.TestValid("yy\\.mm\\.dd");
        this.TestValid("yym");
        this.TestValid("yymm");
        this.TestValid("yymmm");
        this.TestValid("yymmmm");
        this.TestValid("yymmmmm");
        this.TestValid("yyyy");
        this.TestValid("yyyy\"년\"\\ m\"월\"\\ d\"일\";@");
        this.TestValid("yyyy-m-d h:mm AM/PM");
        this.TestValid("yyyy-mm-dd");
        this.TestValid("yyyy/mm/dd");
        this.TestValid("yyyy\\-m\\-d\\ hh:mm:ss");
        this.TestValid("yyyy\\-mm\\-dd");
        this.TestValid("yyyy\\-mm\\-dd;@");
        this.TestValid("yyyy\\-mm\\-dd\\ h:mm");
        this.TestValid("yyyy\\-mm\\-dd\\Thh:mm");
        this.TestValid("yyyy\\-mm\\-dd\\Thhmmss.000");
    }

    [Test]
    [Arguments(null)]
    [Arguments("")]
    [Arguments("General")]
    public void TestDefaultFormatString(string? formatString)
    {
        string? result;

        result = Format(1234.56, formatString!, CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual("1234.56", result);

        result = Format(Double.MaxValue, formatString!, CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual("1.79769313486232E+308", result);

        result = Format(float.MaxValue, formatString!, CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual("3.402823E+38", result);

        result = Format(new DateTime(2017, 10, 28), formatString!, new CultureInfo("sv-se"));
        ClassicAssert.AreEqual("2017-10-28 00:00:00", result);

        result = Format(new DateTime(2017, 10, 28), formatString!, CultureInfo.InvariantCulture);
        ClassicAssert.AreEqual("10/28/2017 00:00:00", result);
    }

    [Test]
    public void TestCurrency()
    {
        this.Test(1234.56, "[$€-1809]# ##0.00", "€1 234.56");
        this.Test(1234.56, "#,##0.00 [$EUR]", "1,234.56 EUR");
    }
}
