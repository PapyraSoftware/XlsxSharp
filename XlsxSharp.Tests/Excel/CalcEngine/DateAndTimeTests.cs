using System;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class DateAndTimeTests
{
    [Test]
    [Arguments(2008, 1, 1, 39448)]
    [Arguments(2008, 15, 1, 39873)]
    [Arguments(2008, -50, 1, 37895)]
    [Arguments(2008, 5, 63, 39631)]
    [Arguments(2008, 13, 63, 39876)]
    [Arguments(2008, 15, -120, 39752)]
    [Arguments(1900, 2, 29, 60)] // Loveable 29th feb 1900
    [Arguments(1900, 2, 28, 59)]
    [Arguments(1900, 1, 1, 1)]
    [Arguments(1900, 1, 0, 0)] // Excel formats it as 1900-01-00, but more like 1899-12-31
    [Arguments(1899, 1, 1, 693598)] // If year < 1900, add 1900 to it
    public void DateReturnsSerialDate(int year, int month, int day, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})").GetNumber()
        );
    }

    [Test]
    [Arguments(1900, 1, -1)] // Serial date -1, below 0
    [Arguments(9999, 12, 32)]
    public void DateReturnsErrorWhenResultOutsideBaseDateToMaxDateOfCalendarSystem(
        int year,
        int month,
        int day
    )
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})");
        ClassicAssert.AreEqual(XLError.NumberInvalid, actual);
    }

    [Test]
    [Arguments(-1, 32000, 1, 973586)] // Year -1.1 behaves as -2
    [Arguments(-1.1, 32000, 1, 973221)]
    [Arguments(-2, 32000, 1, 973221)]
    [Arguments(2000, -5, 1, 36342)] // Month -5.1 behaves as -6
    [Arguments(2000, -5.1, 1, 36312)]
    [Arguments(2000, -6, 1, 36312)]
    [Arguments(2000, 2, -10, 36546)] // Day -10.1 behaves as -11
    [Arguments(2000, 2, -10.1, 36545)]
    [Arguments(2000, 2, -11, 36545)]
    public void DateFloorsArguments(double year, double month, double day, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})").GetNumber()
        );
    }

    [Test]
    [Arguments(10000, -32767, 3, "7269-05-03")] // Month can be [-32767..32767)
    [Arguments(10000, -32767.1, 3, XLError.NumberInvalid)]
    [Arguments(2000, 32766.9, 1, "4730-06-01")]
    [Arguments(2000, 32767, 1, XLError.NumberInvalid)]
    [Arguments(2000, 1, 32767.9, "2089-09-16")] // Day is clamped to at most 32767
    [Arguments(2000, 1, 32768, "2089-09-16")]
    [Arguments(2000, 1, 1E+100, "2089-09-16")]
    [Arguments(2000, 1, -32768, "1910-04-14")] // When day is < -32768, day uses 32767 value instead
    [Arguments(2000, 1, -32768.1, "2089-09-16")]
    [Arguments(2000, 1, -1E+100, "2089-09-16")]
    [Arguments(10000, -32000, 1, "7333-04-01")] // Year is clamped to 10000
    [Arguments(10001, -32000, 1, "7333-04-01")]
    [Arguments(1E+100, -32000, 1, "7333-04-01")]
    [Arguments(-1E+100, 1, 1, XLError.NumberInvalid)] // Even if year is less than int.MinValue, there is no error
    public void DateMatchesExcelBehaviorForOutOfRangeArguments(
        double year,
        double month,
        double day,
        object expectedResult
    )
    {
        if (expectedResult is string iso8601)
        {
            expectedResult = DateTime.Parse(iso8601).ToSerialDateTime();
        }

        XLCellValue actual = XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Arguments("1/1/2006", "12/12/2010", "Y", 4)]
    [Arguments("1/1/2006", "12/12/2010", "M", 59)]
    [Arguments("1/1/2006", "12/12/2010", "D", 1806)]
    [Arguments("1/1/2006", "12/12/2010", "MD", 11)]
    [Arguments("1/1/2006", "12/12/2010", "YM", 11)]
    [Arguments("1/1/2006", "12/12/2010", "YD", 345)]
    [Arguments(38718, 40524, "Y", 4)]
    [Arguments(38718, 40524, "M", 59)]
    [Arguments(38718, 40524, "D", 1806)]
    [Arguments(38718, 40524, "MD", 11)]
    [Arguments("2020-01-31", "2024-03-01", "MD", -1)] // Pathological case. Start is shifted to 2024-02-31, thus 2024-03-02 is one day before the end
    [Arguments("1990-01-20", "2002-12-15", "YM", 10)] // YM across many years
    [Arguments(38718, 40524, "YM", 11)]
    [Arguments(38718, 40524, "YD", 345)]
    [Arguments("2001-12-31", "2002-4-15", "YM", 3)] // YM counts only full months - the last month is not full
    [Arguments("2001-12-10", "2002-4-15", "YM", 4)] // YM counts only full months - the last month is full
    [Arguments("2001-12-15", "2002-4-15", "YM", 4)] // YM counts only full months - the last month exactly full
    [Arguments("1900-01-12", "1901-03-04", "YD", 51)] // YD has plus +1 error with start dates in jan/feb 1900 and end in march of subsequent years
    [Arguments("2001-12-31", "2002-4-15", "YD", 105)] // YD ignores year, baseline
    [Arguments("2001-12-31", "2003-4-15", "YD", 105)] // YD ignores year, different year
    [Arguments("2000-02-20", "2100-02-10", "YD", 356)] // YD uses start year, not end year. Start has feb29, baseline
    [Arguments("2001-02-20", "2100-02-10", "YD", 355)] // YD uses start year, not end year. Start doesn't have feb29 => it's one less than the baseline
    [Arguments("2002-01-31", "2002-4-15", "YD", 74)]
    [Arguments("2001-12-02", "2001-12-15", "Y", 0)]
    [Arguments("2001-12-02", "2002-12-02", "Y", 1)]
    [Arguments("2006-01-15", "2006-03-14", "M", 1)]
    [Arguments("2020-11-22", "2020-11-23 9:00", "D", 1)]
    public void DateDif(object startDate, object endDate, string unit, double expected)
    {
        if (startDate is string s1)
        {
            startDate = $"\"{s1}\"";
        }

        if (endDate is string s2)
        {
            endDate = $"\"{s2}\"";
        }

        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"DATEDIF({startDate},{endDate},\"{unit}\")")
        );
    }

    [Test]
    [Arguments("N")]
    public void DateDifReturnsNumberErrorOnUnexpectedUnit(string unit) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"DATEDIF(10,100,\"{unit}\")")
        );

    [Test]
    public void DateDifEndDateCantBeAfterStartDate() =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("DATEDIF(40524,38718,\"D\")")
        );

    [Test]
    [Arguments(-0.1, 100)]
    [Arguments(1, 2958466)]
    public void DateDifReturnsNumberErrorOnDateOutOfDateSystem(
        decimal startDate,
        decimal endDate
    ) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"DATEDIF({startDate},{endDate},\"D\")")
        );

    [Test]
    [Arguments("8/22/2008", 39682)]
    [Arguments("2/1/2006", 38749)]
    [Arguments("2006-2-1", 38749)]
    [Arguments("22-MAY-2011", 40685)]
    [Arguments("February 1, 2006 17:45", 38749)]
    public void DateValueReturnsTruncatedSerialDateExtractedFromText(string date, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExprCurrent($"DATEVALUE(\"{date}\")")
        );
    }

    [Test]
    public void DateValueReturnsTruncatedSerialDateUsingCurrentYear()
    {
        // If year isn't provided in string, it should parse as "current year"
        double actual = (double)XLWorkbook.EvaluateExpr("DATEVALUE(\"5-JUL\")");
        double expected = new DateTime(DateTime.Now.Year, 7, 5).ToOADate();
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Arguments("\"100\"")]
    [Arguments("\"0\"")]
    public void DateValueDoesntCoerceNumberInATextToADate(string arg) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExprCurrent($"DATEVALUE({arg})")
        );

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("1000")]
    [Arguments("DATE(2006,1,5)")]
    public void DateValueReturnsCoercionErrorOnNonText(string arg) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExprCurrent($"DATEVALUE({arg})")
        );

    [Test]
    public void DateValuePropagatesError() =>
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExprCurrent("DATEVALUE(#DIV/0!)")
        );

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.5, 0)]
    [Arguments(1, 1)]
    [Arguments(31, 31)]
    [Arguments(32, 1)]
    [Arguments(59, 28)]
    [Arguments(60, 29)]
    [Arguments(61, 1)]
    [Arguments(30000, 18)]
    [Arguments(45718, 2)]
    public void DayReturnsDayOfAMonthForSerialCulture(double serialDate, double expected)
    {
        ClassicAssert.AreEqual(expected, XLWorkbook.EvaluateExpr($"DAY({serialDate})").GetNumber());
    }

    [Test]
    [Arguments("\"8/22/2008\"", 22)]
    [Arguments("\"1/2/2006 10:45 AM\"", 2)]
    [Arguments("\"367\"", 1)]
    [Arguments("IF(TRUE,)", 0)] // Blank
    [Arguments("TRUE", 1)]
    [Arguments("FALSE", 0)]
    public void DayAcceptsNonNumberValues(string value, double expected)
    {
        ClassicAssert.AreEqual(expected, XLWorkbook.EvaluateExpr($"DAY({value})").GetNumber());
    }

    [Test]
    [Skip("Excel accepts this but XlsxSharp does not yet")]
    public void DayAcceptsMissingYearAndSubstitutesCurrentYear()
    {
        // Test providing just month and day, which should fill the year as "current year"
        double actual = XLWorkbook.EvaluateExpr("DAY(\"8/22\")").GetNumber();
        ClassicAssert.AreEqual(22, actual);
    }

    [Test]
    public void DayOnlyAcceptsSerialDateFrom0ToUpperLimitOfCalendarSystem()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAY(-0.1)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("DAY(DATE(9999,12,31)+1)")
        );
    }

    [Test]
    [Culture("eu-ES")]
    [Arguments("\"2006/1/2 10:45 AM\"", 2)]
    [Arguments("DATE(2006,1,2)", 2)]
    [Arguments("DATE(2006,0,2)", 2)]
    [Arguments("DATE(2013,9,0)", 31)]
    public void DayExamples(string date, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExprCurrent($"DAY({date})").GetNumber()
        );
    }

    [Test]
    [Arguments(2016, 10, 1, 1992, 2, 29, 8981)]
    [Arguments(1901, 3, 10, 1900, 1, 26, 409)]
    public void DaysCalculateDifferenceBetweenTwoDates(
        double endYear,
        double endMonth,
        double endDay,
        double startYear,
        double startMonth,
        double startDay,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)
                XLWorkbook.EvaluateExpr(
                    $"DAYS(DATE({endYear},{endMonth},{endDay}),DATE({startYear},{startMonth},{startDay}))"
                )
        );
    }

    [Test]
    [Arguments("2016-10-01", "1992-02-29", 8981)]
    [Arguments("1901-03-10", "1900-01-26", 409)]
    [Arguments("1900-01-26", "1901-03-10", -409)]
    public void DaysCoercesDatesToNumber(string endDate, string startDate, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"DAYS(\"{endDate}\",\"{startDate}\")")
        );
    }

    [Test]
    public void DaysTruncatesPassedArguments() =>
        ClassicAssert.AreEqual(9, XLWorkbook.EvaluateExpr("DAYS(10.6,1.9)"));

    [Test]
    public void DaysArgumentsMustBeInDateRange()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAYS(-0.1,1)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAYS(2958466,1)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAYS(1,-0.1)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAYS(1,2958466)"));
    }

    [Test]
    public void Days360UsesUsMethodByDefault()
    {
        const string formulaFormat = "DAYS360(DATE(2002,2,3),DATE(2005,5,31){0})";
        XLCellValue defaultResult = XLWorkbook.EvaluateExpr(
            string.Format(formulaFormat, string.Empty)
        );
        XLCellValue usResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, ",FALSE"));
        XLCellValue euResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, ",TRUE"));
        ClassicAssert.AreEqual(1198, defaultResult);
        ClassicAssert.AreEqual(usResult, defaultResult);
        ClassicAssert.AreNotEqual(euResult, defaultResult);
    }

    [Test]
    public void Days360Europe1()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("DAYS360(\"1/1/2008\", \"3/31/2008\",TRUE)");
        ClassicAssert.AreEqual(89, actual);
    }

    [Test]
    public void Days360Europe2()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("DAYS360(\"3/31/2008\", \"1/1/2008\",TRUE)");
        ClassicAssert.AreEqual(-89, actual);
    }

    [Test]
    [Arguments(2002, 2, 3, 2005, 5, 31, 1198)]
    [Arguments(2005, 5, 31, 2002, 2, 3, -1197)]
    [Arguments(2008, 1, 1, 2008, 3, 31, 90)]
    [Arguments(2008, 3, 31, 2008, 1, 1, -89)]
    [Arguments(2020, 2, 29, 2021, 2, 28, 358)]
    [Arguments(2020, 5, 29, 2020, 4, 1, -58)]
    [Arguments(2020, 5, 29, 2020, 3, 31, -58)]
    [Arguments(2020, 5, 30, 2020, 4, 1, -59)]
    [Arguments(2020, 5, 30, 2020, 3, 31, -60)]
    [Arguments(2020, 5, 30, 2020, 3, 30, -60)]
    public void Days360UsMethod(
        int startYear,
        int startMonth,
        int startDay,
        int endYear,
        int endMonth,
        int endDay,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)
                XLWorkbook.EvaluateExpr(
                    $"DAYS360(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),FALSE)"
                )
        );
    }

    [Test]
    [Arguments(1900, 2, 27, 1900, 2, 27, 0)]
    [Arguments(1900, 2, 27, 1900, 2, 28, 1)]
    [Arguments(1900, 2, 27, 1900, 2, 29, 2)]
    [Arguments(1900, 2, 27, 1900, 3, 1, 4)]
    [Arguments(1900, 2, 28, 1900, 2, 27, -1)]
    [Arguments(1900, 2, 28, 1900, 2, 28, 0)]
    [Arguments(1900, 2, 28, 1900, 2, 29, 1)]
    [Arguments(1900, 2, 28, 1900, 3, 1, 3)]
    [Arguments(1900, 2, 29, 1900, 2, 27, -3)]
    [Arguments(1900, 2, 29, 1900, 2, 28, -2)]
    [Arguments(1900, 2, 29, 1900, 2, 29, -1)]
    [Arguments(1900, 2, 29, 1900, 3, 1, 1)]
    [Arguments(1900, 3, 1, 1900, 2, 27, -4)]
    [Arguments(1900, 3, 1, 1900, 2, 28, -3)]
    [Arguments(1900, 3, 1, 1900, 2, 29, -2)]
    [Arguments(1900, 3, 1, 1900, 3, 1, 0)]
    public void Days360UsMethodForFeb291900(
        int startYear,
        int startMonth,
        int startDay,
        int endYear,
        int endMonth,
        int endDay,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)
                XLWorkbook.EvaluateExpr(
                    $"DAYS360(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),FALSE)"
                )
        );
    }

    [Test]
    [Arguments("2008-03-01", -1, "2008-02-01")]
    [Arguments("2008-03-31", -1, "2008-02-29")]
    [Arguments("2008-03-01", 1, "2008-04-01")]
    [Arguments("2008-03-31", 1, "2008-04-30")]
    [Arguments("2008-03-01", -1, "2008-02-01")]
    [Arguments("2008-03-31", 1, "2008-04-30")]
    [Arguments("1900-01-31", 1, "1900-02-28")] // Uses correct FEB28
    [Arguments("1900-01-31", 2, "1900-03-31")]
    [Arguments("1983-07-31", -77, "1977-02-28")]
    [Arguments("2021-05-14", 35, "2024-04-14")]
    public void EDateReturnsEndDateFromStartDateAndMonthOffset(
        string startDate,
        double monthOffset,
        string expectedEndDate
    )
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"EDATE(\"{startDate}\",{monthOffset})");
        ClassicAssert.AreEqual(DateTime.Parse(expectedEndDate).ToSerialDateTime(), actual);
    }

    [Test]
    public void EDateReturnsNumberErrorForNonDateValues()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("EDATE(-0.1,0)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("EDATE(2958466,0)"));
    }

    [Test]
    [Arguments("1900-01-01", -1)]
    [Arguments("9999-07-10", 6)]
    [Arguments("9999-07-10", 1E+100)]
    public void EDateReturnsNumberErrorWhenEndDateIsOutOfDateSystem(
        string startDate,
        double monthOffset
    ) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"EDATE(\"{startDate}\",{monthOffset})")
        );

    [Test]
    [Arguments(1900, 1, 0, 0, 31)]
    [Arguments(1900, 1, 1, 0, 31)]
    [Arguments(1900, 1, 31, 0, 31)]
    [Arguments(1900, 2, 20, 0, 59)]
    [Arguments(1900, 2, 29, 0, 59)]
    [Arguments(1900, 2, 29, 1, 91)]
    [Arguments(1900, 2, 29, 1, 91)]
    [Arguments(1900, 3, 1, -1, 59)]
    [Arguments(1985, 4, 15, 9, 31443)]
    [Arguments(2006, 1, 31, 5, 38898)] // Spec examples
    [Arguments(2004, 2, 29, 12, 38411)]
    [Arguments(2004, 2, 28, 12, 38411)]
    [Arguments(2004, 1, 15, -23, 37315)]
    public void EomonthReturnsEndOfMonthFromStartDatePlusMonthOffset(
        int year,
        int month,
        int day,
        int months,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"EOMONTH(DATE({year},{month},{day}),{months})")
        );
    }

    [Test]
    public void EomonthTruncatesArguments() =>
        ClassicAssert.AreEqual(59, XLWorkbook.EvaluateExpr("EOMONTH(60.1,0.9)"));

    [Test]
    public void EomonthStartDateMustBeInDateValues()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("EOMONTH(-0.1,0)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("EOMONTH(DATE(9999,12,31)+1,0)")
        );
    }

    [Test]
    [Arguments("1900-01-01", -1)]
    [Arguments("9999-12-10", 1)]
    public void EomonthReturnsNumberErrorWhenEndDateIsOutOfDateSystem(
        string startDate,
        double monthOffset
    ) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"EOMONTH(\"{startDate}\",{monthOffset})")
        );

    [Test]
    [Arguments("0", 0)]
    [Arguments("0.25", 6)]
    [Arguments("0.5", 12)]
    [Arguments("0.75", 18)]
    [Arguments("1", 0)]
    [Arguments("1.75", 18)]
    [Arguments("\"1.75\"", 18)] // Test string in addition to number in TestCase before
    [Arguments("\"7/18/2011 7:45\"", 7)]
    [Arguments("\"4/21/2012\"", 0)]
    [Arguments("\"12:00:00\"", 12)]
    [Arguments("\"8/22/2008 3:30:45 PM\"", 15, Skip = "We don't parse seconds")]
    [Arguments("\"8/22/2008 3:30 PM\"", 15)]
    [Arguments("DATE(2006,2,26)+TIME(2,10,20)", 2)]
    [Arguments("TIME(22,56,34)", 22)]
    [Arguments(
        "\"22-Oct-2001 10:53:12\"",
        10,
        Skip = "We don't parse seconds plus culture is wrong"
    )]
    [Arguments("\"October 22, 2001 10:53\"", 10)]
    [Arguments("\"10:53:12 pm\"", 22)]
    [Arguments("\"22:53:12\"", 22)]
    [Arguments("IF(TRUE,)", 0)] // Blank
    [Arguments("TRUE", 0)]
    [Arguments("FALSE", 0)]
    public void HourReturnsHourOfSerialDate(string dateArg, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExprCurrent($"HOUR({dateArg})").GetNumber()
        );
    }

    [Test]
    public void HourAcceptsOnlySerialTimeBetweenZeroAndUpperLimitOfDateSystem()
    {
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExprCurrent("HOUR(0)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("HOUR(-0.1)"));

        ClassicAssert.AreEqual(21, XLWorkbook.EvaluateExprCurrent("HOUR(DATE(9999,12,31)+0.9)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExprCurrent("HOUR(DATE(9999,12,31)+1)")
        );
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("0.5", 0)]
    [Arguments("0.68", 19)]
    [Arguments("0.69", 33)]
    [Arguments("0.85", 24)]
    [Arguments("10.85", 24)]
    [Arguments("\"10.85\"", 24)] // Test string in addition to number in TestCase before
    [Arguments("\"14:47:20\"", 47)]
    [Arguments("\"8/22/2008 3:30 AM\"", 30)]
    [Arguments("IF(TRUE,)", 0)] // Blank
    [Arguments("TRUE", 0)]
    [Arguments("FALSE", 0)]
    public void MinuteReturnsMinuteOfSerialDate(string dateArg, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExprCurrent($"MINUTE({dateArg})").GetNumber()
        );
    }

    [Test]
    public void MinuteAcceptsOnlySerialTimeBetweenZeroAndUpperLimitOfDateSystem()
    {
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExprCurrent("MINUTE(0)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExprCurrent("MINUTE(-0.1)")
        );

        ClassicAssert.AreEqual(36, XLWorkbook.EvaluateExprCurrent("MINUTE(DATE(9999,12,31)+0.9)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExprCurrent("MINUTE(DATE(9999,12,31)+1)")
        );
    }

    [Test]
    [Culture("eu-ES")]
    [Arguments(0, 1)] // 1900-01-00
    [Arguments(31, 1)] // 1900-01-31
    [Arguments(32, 2)] // 1900-02-01
    [Arguments(59, 2)] // 1900-02-28
    [Arguments(60, 2)] // 1900-02-29
    [Arguments(61, 3)] // 1900-03-01
    [Arguments("DATE(2006,1,2)", 1)]
    [Arguments("DATE(2006,0,2)", 12)]
    [Arguments("\"2006/1/2 10:45 AM\"", 1)]
    [Arguments("30000", 2)]
    [Arguments("45596", 10)]
    [Arguments("45596.9", 10)]
    [Arguments("45597", 11)]
    [Arguments("\"45597\"", 11)] // Test string in addition to number in TestCase before
    [Arguments("IF(TRUE,)", 1)] // Blank
    [Arguments("TRUE", 1)]
    [Arguments("FALSE", 1)]
    public void MonthReturnsMonthOfSerialDate(object argument, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExprCurrent($"MONTH({argument})").GetNumber()
        );
    }

    [Test]
    [Skip("Excel accepts this but XlsxSharp does not yet")]
    public void MonthAcceptsMissingYearAndSubstitutesCurrentYear()
    {
        // Test providing just month and day, which should fill the year as "current year"
        double actual = XLWorkbook.EvaluateExpr("MONTH(\"8/22\")").GetNumber();
        ClassicAssert.AreEqual(8, actual);
    }

    [Test]
    public void MonthSerialDateMustBeBetweenZeroAndUpperLimitOfDateSystem()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("MONTH(-0.1)"));
        ClassicAssert.AreEqual(12, XLWorkbook.EvaluateExpr("MONTH(DATE(9999,12,31) + 0.9)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("MONTH(DATE(9999,12,31) + 1)")
        );
    }

    [Test]
    [Arguments(1900, 1, 0, 52)]
    [Arguments(1900, 1, 1, 52)]
    [Arguments(1900, 1, 2, 1)]
    [Arguments(1900, 2, 28, 9)]
    [Arguments(1900, 2, 29, 9)]
    [Arguments(1900, 3, 1, 9)]
    [Arguments(2012, 1, 2, 1)]
    [Arguments(2012, 12, 31, 1)]
    [Arguments(2012, 3, 9, 10)]
    [Arguments(2014, 12, 12, 50)]
    [Arguments(9999, 12, 31, 52)]
    public void IsoWeekNum(int year, int month, int day, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"ISOWEEKNUM(DATE({year},{month},{day}))")
        );
    }

    [Test]
    public void NetWorkDaysWithHolidays()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell()
            .SetValue("Date")
            .CellBelow()
            .SetValue(new DateTime(2008, 10, 1))
            .CellBelow()
            .SetValue(new DateTime(2009, 3, 1))
            .CellBelow()
            .SetValue(new DateTime(2008, 11, 26))
            .CellBelow()
            .SetValue(new DateTime(2008, 12, 4))
            .CellBelow()
            .SetValue(new DateTime(2009, 1, 21))
            .CellBelow()
            .SetValue(new DateTime(2009, 1, 4)) // Holiday is on Sunday - do not count twice
            .CellBelow()
            .SetValue(new DateTime(2009, 1, 6)) // Workweek holiday is specified twice, shouldn't be counted twice
            .CellBelow()
            .SetValue(new DateTime(2009, 1, 6))
            .CellBelow()
            .SetValue(new DateTime(2008, 9, 30)) // Tuesday holiday just before the first date, shouldn't be counted
            .CellBelow()
            .SetValue(
                new DateTime(2009, 3, 2)
            ) // Monday holiday just after the last date, shouldn't be counted
        ;
        XLCellValue actual = ws.Evaluate("NETWORKDAYS(A2, A3, A4:A11)");
        ClassicAssert.AreEqual(104, actual);
    }

    [Test]
    [Arguments("2024-10-01", "2024-10-01", 1)] // Tue-Tue
    [Arguments("2024-10-01", "2024-10-02", 2)] // Tue-Wed
    [Arguments("2024-10-01", "2024-10-03", 3)] // Tue-Thu
    [Arguments("2024-10-01", "2024-10-04", 4)] // Tue-Fri
    [Arguments("2024-10-01", "2024-10-05", 4)] // Tue-Sat
    [Arguments("2024-10-01", "2024-10-06", 4)] // Tue-Sun
    [Arguments("2024-10-01", "2024-10-07", 5)] // Tue-Mon
    [Arguments("2024-09-29", "2024-10-12", 10)] // Sun-Sat
    [Arguments("2024-09-29", "2024-10-13", 10)] // Sun-Sun
    [Arguments("2024-09-29", "2024-10-14", 11)] // Sun-Mon
    [Arguments("2024-09-29", "2024-10-15", 12)] // Sun-Tue
    [Arguments("2024-09-29", "2024-10-16", 13)] // Sun-Wed
    [Arguments("2024-09-29", "2024-10-17", 14)] // Sun-Thu
    [Arguments("2024-09-29", "2024-10-18", 15)] // Sun-Fri
    [Arguments("2024-09-29", "2024-10-19", 15)] // Sun-Sat
    public void NetWorkDaysNonFullWeeksAreCountedCorrectly(
        string startDate,
        string endDate,
        int expected
    )
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            $"NETWORKDAYS(\"{startDate}\", \"{endDate}\")"
        );
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void NetWorkDaysWithEndDateEarlierThanStartDate()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("NETWORKDAYS(\"3/01/2009\", \"10/01/2008\")");
        ClassicAssert.AreEqual(-108, actual);

        actual = XLWorkbook.EvaluateExpr("NETWORKDAYS(\"2016-01-01\", \"2015-12-23\")");
        ClassicAssert.AreEqual(-8, actual);
    }

    [Test]
    public void NetWorkDaysBehavior()
    {
        using XLWorkbook wb = new();
        XLCellValue actual = wb.Evaluate(
            "NETWORKDAYS(\"10/01/2008\", \"3/01/2009\", \"11/26/2008\")"
        );
        ClassicAssert.AreEqual(107, actual);

        // Example from specification. Except spec wrong. The value is 1 off from Excel value.
        ClassicAssert.AreEqual(22, wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 1), DATE(2006, 1, 31))"));
        ClassicAssert.AreEqual(
            -22,
            wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 31), DATE(2006, 1, 1))")
        );
        ClassicAssert.AreEqual(
            21,
            wb.Evaluate(
                "NETWORKDAYS(DATE(2006, 1, 1), DATE(2006, 2, 1), { \"2006-01-02\", \"2006-01-16\" })"
            )
        );

        // Scalar number is accepted for holidays
        ClassicAssert.AreEqual(6, wb.Evaluate("NETWORKDAYS(1, 10, 2)"));

        // Scalar logical causes conversion error
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(TRUE, 10)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(0, TRUE)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, 10, TRUE)"));

        // Scalar text is converted
        ClassicAssert.AreEqual(6, wb.Evaluate("NETWORKDAYS(\"1\", \"10\", \"2\")"));
        ClassicAssert.AreEqual(6, wb.Evaluate("NETWORKDAYS(1, 10, \"0 4/2\")"));
        ClassicAssert.AreEqual(6, wb.Evaluate("NETWORKDAYS(1, 10, \"1900-01-02\")"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(\"Text\", 10)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, \"Text\")"));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            wb.Evaluate("NETWORKDAYS(1, 10, \"Text\")")
        );

        // Array accepts numbers and converts text
        ClassicAssert.AreEqual(5, wb.Evaluate("NETWORKDAYS(1, 10, {\"2\", 3})"));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            wb.Evaluate("NETWORKDAYS(1, 10, {\"Text\"})")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            wb.Evaluate("NETWORKDAYS(1, 10, {TRUE})")
        );

        // Same conversion logic applies to reference values
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value; // Ignored
        ws.Cell("A2").Value = false; // Causes conversion error
        ws.Cell("A3").Value = true; // Causes conversion error
        ws.Cell("A4").Value = 37147; // 2001-09-13
        ws.Cell("A5").Value = "2001-09-12"; // Monday
        ws.Cell("A6").Value = XLError.NoValueAvailable;

        ClassicAssert.AreEqual(175, ws.Evaluate("NETWORKDAYS(\"2001-05-01\", \"2001-12-31\", A1)"));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            ws.Evaluate("NETWORKDAYS(\"2001-05-01\", \"2001-12-31\", A1:A3)")
        );
        ClassicAssert.AreEqual(
            173,
            ws.Evaluate("NETWORKDAYS(\"2001-05-01\",\"2001-12-31\", A4:A5)")
        );

        // Errors are propagated
        ClassicAssert.AreEqual(XLError.NoValueAvailable, wb.Evaluate("NETWORKDAYS(#N/A, 10)"));
        ClassicAssert.AreEqual(XLError.NoValueAvailable, wb.Evaluate("NETWORKDAYS(1, #N/A)"));
        ClassicAssert.AreEqual(XLError.NoValueAvailable, wb.Evaluate("NETWORKDAYS(1, 10, {#N/A})"));
        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("NETWORKDAYS(1, 10, A6)"));
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("0.5", 0)]
    [Arguments("1", 0)]
    [Arguments("366", 0)]
    [Arguments("367", 0)]
    [Arguments("\"367\"", 0)] // Test string in addition to number in TestCase before
    [Arguments("\"8/22/2008\"", 0)]
    [Arguments("\"1/2/2006 10:45 AM\"", 0)]
    [Arguments("\"8/22/2008 3:30:4 PM\"", 4, Skip = "We don't parse seconds")]
    [Arguments("\"8/22/2008 3:30:23 PM\"", 23, Skip = "We don't parse seconds")]
    [Arguments("\"3:30:45\"", 45)]
    [Arguments("IF(TRUE,)", 0)] // Blank
    [Arguments("TRUE", 0)]
    [Arguments("FALSE", 0)]
    public void SecondReturnsSecondOfSerialDate(string dateArg, double expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExprCurrent($"SECOND({dateArg})").GetNumber()
        );
    }

    [Test]
    public void SecondAcceptsOnlySerialTimeBetweenZeroAndUpperLimitOfDateSystem()
    {
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExprCurrent("SECOND(0)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExprCurrent("SECOND(-0.1)")
        );

        ClassicAssert.AreEqual(
            51,
            XLWorkbook.EvaluateExprCurrent("SECOND(DATE(9999,12,31)+0.9999)")
        );
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExprCurrent("SECOND(DATE(9999,12,31)+1)")
        );
    }

    [Test]
    [Arguments(0, 0, 0, 0)]
    [Arguments(0, 0, 1, 0.0000115740740741)]
    [Arguments(0, 0, 2, 0.0000231481481481)]
    [Arguments(0, 0, 20, 0.0002314814814815)]
    [Arguments(2, 3, 20, 0.0856481481481481)]
    [Arguments(12, 0, 0, 0.5000000000000000)]
    [Arguments(23, 59, 59, 0.9999884259259260)]
    [Arguments(26, 120, 240, 0.1694444444444450)]
    [Arguments(1, 2, 3, 0.043090277777778)]
    [Arguments(1.9, 2.9, 3.9, 0.043090277777778)]
    [Arguments(24, 0, 0, 0)]
    [Arguments(0, 42 * 60, 0, 0.75)]
    [Arguments(0, 0, 60 * 60 * 3, 0.125)]
    [Arguments(120, 240, 347, 0.170682870370)]
    public void TimeReturnsSerialDateTime(
        double hour,
        double minute,
        double second,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExpr($"TIME({hour},{minute},{second})"),
            XLHelper.Epsilon
        );
    }

    [Test]
    [Arguments(-0.1, 0, 0)]
    [Arguments(32768, 0, 0)]
    [Arguments(0, -0.1, 0)]
    [Arguments(0, 32768, 0)]
    [Arguments(0, 0, -0.1)]
    [Arguments(0, 0, 32768)]
    public void TimeComponentsMustBeInZeroTo32767Interval(
        double hour,
        double minute,
        double second
    ) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"TIME({hour},{minute},{second})")
        );

    [Test]
    [Arguments("2:24 AM", 0.1)]
    [Arguments("August 22, 2008 6:35 AM", 0.27430555555555558)]
    public void TimeValueReturnsTimeComponentOfSerialDateExtractedFromText(
        string time,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)XLWorkbook.EvaluateExprCurrent($"TIMEVALUE(\"{time}\")"),
            XLHelper.Epsilon
        );
    }

    [Test]
    [Arguments("\"10.5\"")]
    [Arguments("\"0\"")]
    public void TimeValueDoesntCoerceNumberInATextToATime(string numberText) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExprCurrent($"TIMEVALUE({numberText})")
        );

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("0.25")]
    [Arguments("TIME(18,25,48)")]
    public void TimeValueReturnsCoercionErrorOnNonText(string nonText) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExprCurrent($"TIMEVALUE({nonText})")
        );

    [Test]
    public void TimeValuePropagatesError() =>
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExprCurrent("TIMEVALUE(#DIV/0!)")
        );

    [Test]
    public void Today()
    {
        double actual = (double)XLWorkbook.EvaluateExpr("TODAY()");
        ClassicAssert.AreEqual(DateTime.Today.ToSerialDateTime(), actual);
    }

    [Test]
    [Arguments("\"2/14/2008\"", 1, 5)]
    [Arguments("\"2/14/2008\"", 2, 4)]
    [Arguments("\"2/14/2008\"", 3, 3)]
    [Arguments("\"2/14/2008\"", 11, 4)]
    [Arguments("\"2/14/2008\"", 12, 3)]
    [Arguments("\"2/14/2008\"", 13, 2)]
    [Arguments("\"2/14/2008\"", 14, 1)]
    [Arguments("\"2/14/2008\"", 15, 7)]
    [Arguments("\"2/14/2008\"", 16, 6)]
    [Arguments("\"2/14/2008\"", 17, 5)]
    public void WeekdayCalculatesWeekDay(string value, int flag, int expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"WEEKDAY({value}, {flag})");
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void WeekdayWithoutFlag()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("WEEKDAY(\"2/14/2008\")");
        ClassicAssert.AreEqual(5, actual);
    }

    [Test]
    public void WeekdayBehavior()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = 45577;
        ClassicAssert.AreEqual(7, ws.Evaluate("WEEKDAY(A1)"));

        // Time of the day doesn't matter, serial date is truncated
        ClassicAssert.AreEqual(7, XLWorkbook.EvaluateExpr("WEEKDAY(45577.9, 1.9)"));

        ClassicAssert.AreEqual(7, XLWorkbook.EvaluateExpr("WEEKDAY(0)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(-1)"));

        // Year 10k
        ClassicAssert.AreEqual(6, XLWorkbook.EvaluateExpr("WEEKDAY(2958465)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(2958466)"));

        // Convert from logical/text to number
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(TRUE)"));
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(\"0 2/2\")"));
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(1, TRUE)"));
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(1, \"1 0/2\")"));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("WEEKDAY(\"text\")")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("WEEKDAY(1, \"text\")")
        );

        // Flag can only have some values
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 0)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 4)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 10)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 18)"));

        // Error is propagated
        ClassicAssert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr("WEEKDAY(#N/A)"));
        ClassicAssert.AreEqual(
            XLError.NoValueAvailable,
            XLWorkbook.EvaluateExpr("WEEKDAY(5, #N/A)")
        );
    }

    [Test]
    [Arguments(1, 1986, 12, 27, 52)]
    [Arguments(1, 1986, 12, 28, 53)]
    [Arguments(1, 1986, 12, 31, 53)]
    [Arguments(1, 1987, 1, 1, 1)]
    [Arguments(1, 1987, 1, 3, 1)]
    [Arguments(1, 1987, 1, 4, 2)]
    [Arguments(1, 2000, 3, 9, 11)]
    [Arguments(1, 2002, 3, 9, 10)]
    [Arguments(1, 2003, 3, 9, 11)]
    [Arguments(1, 2004, 3, 9, 11)]
    [Arguments(1, 2005, 3, 9, 11)]
    [Arguments(1, 2006, 3, 9, 10)]
    [Arguments(1, 2007, 3, 9, 10)]
    [Arguments(1, 2008, 3, 9, 11)]
    [Arguments(1, 2009, 3, 9, 11)]
    [Arguments(2, 1988, 12, 25, 52)]
    [Arguments(2, 1988, 12, 26, 53)]
    [Arguments(2, 1988, 12, 31, 53)]
    [Arguments(2, 1989, 1, 1, 1)]
    [Arguments(2, 1989, 1, 2, 2)]
    [Arguments(2, 2000, 3, 9, 11)]
    [Arguments(2, 2001, 3, 9, 10)]
    [Arguments(2, 2002, 3, 9, 10)]
    [Arguments(2, 2003, 3, 9, 10)]
    [Arguments(2, 2004, 3, 9, 11)]
    [Arguments(2, 2005, 3, 9, 11)]
    [Arguments(2, 2006, 3, 9, 11)]
    [Arguments(2, 2007, 3, 9, 10)]
    [Arguments(2, 2008, 3, 9, 10)]
    [Arguments(2, 2009, 3, 9, 11)]
    [Arguments(11, 1990, 12, 23, 51)]
    [Arguments(11, 1990, 12, 24, 52)]
    [Arguments(11, 1990, 12, 30, 52)]
    [Arguments(11, 1990, 12, 31, 53)]
    [Arguments(11, 1991, 1, 1, 1)]
    [Arguments(11, 1991, 1, 6, 1)]
    [Arguments(11, 1991, 1, 7, 2)]
    [Arguments(12, 1992, 12, 28, 52)]
    [Arguments(12, 1992, 12, 29, 53)]
    [Arguments(12, 1992, 12, 31, 53)]
    [Arguments(12, 1993, 1, 1, 1)]
    [Arguments(12, 1993, 1, 4, 1)]
    [Arguments(12, 1993, 1, 5, 2)]
    [Arguments(13, 1994, 12, 27, 52)]
    [Arguments(13, 1994, 12, 28, 53)]
    [Arguments(13, 1994, 12, 31, 53)]
    [Arguments(13, 1995, 1, 1, 1)]
    [Arguments(13, 1995, 1, 3, 1)]
    [Arguments(13, 1995, 1, 4, 2)]
    [Arguments(14, 1999, 12, 29, 52)]
    [Arguments(14, 1999, 12, 30, 53)]
    [Arguments(14, 1999, 12, 31, 53)]
    [Arguments(14, 2000, 1, 1, 1)]
    [Arguments(14, 2000, 1, 5, 1)]
    [Arguments(14, 2000, 1, 6, 2)]
    [Arguments(15, 2004, 12, 24, 53)]
    [Arguments(15, 2004, 12, 30, 53)]
    [Arguments(15, 2004, 12, 31, 54)]
    [Arguments(15, 2005, 1, 1, 1)]
    [Arguments(15, 2005, 1, 6, 1)]
    [Arguments(15, 2005, 1, 7, 2)]
    [Arguments(16, 2008, 12, 26, 52)]
    [Arguments(16, 2008, 12, 27, 53)]
    [Arguments(16, 2008, 12, 31, 53)]
    [Arguments(16, 2009, 1, 1, 1)]
    [Arguments(16, 2009, 1, 2, 1)]
    [Arguments(16, 2009, 1, 3, 2)]
    [Arguments(16, 2009, 1, 9, 2)]
    [Arguments(17, 1929, 12, 21, 51)]
    [Arguments(17, 1929, 12, 22, 52)]
    [Arguments(17, 1929, 12, 28, 52)]
    [Arguments(17, 1929, 12, 29, 53)]
    [Arguments(17, 1929, 12, 31, 53)]
    [Arguments(17, 1930, 1, 1, 1)]
    [Arguments(17, 1930, 1, 4, 1)]
    [Arguments(17, 1930, 1, 5, 2)]
    [Arguments(17, 1930, 1, 11, 2)]
    [Arguments(21, 1964, 12, 27, 52)]
    [Arguments(21, 1964, 12, 28, 53)]
    [Arguments(21, 1964, 12, 31, 53)]
    [Arguments(21, 1965, 1, 1, 53)]
    [Arguments(21, 1965, 1, 3, 53)]
    [Arguments(21, 1965, 1, 4, 1)]
    [Arguments(21, 1968, 12, 29, 52)]
    [Arguments(21, 1968, 12, 30, 1)]
    [Arguments(21, 1968, 12, 31, 1)]
    [Arguments(21, 1969, 1, 1, 1)]
    [Arguments(21, 1969, 1, 5, 1)]
    [Arguments(21, 1969, 1, 6, 2)]
    public void WeeknumReturnsWeekNumberForDate(
        double weekStartFlag,
        double year,
        double month,
        double day,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook
                .EvaluateExpr($"WEEKNUM(DATE({year},{month},{day}),{weekStartFlag})")
                .GetNumber()
        );
    }

    [Test]
    public void WeeknumDefaultWeekStartsOnSunday()
    {
        for (int day = 14; day <= 20; day++)
        {
            XLCellValue defaultValue = XLWorkbook.EvaluateExpr($"WEEKNUM(DATE(1967,5,{day}))");
            XLCellValue sundayValue = XLWorkbook.EvaluateExpr($"WEEKNUM(DATE(1967,5,{day}),1)");
            ClassicAssert.AreEqual(sundayValue, defaultValue);
        }
    }

    [Test]
    [Arguments]
    public void WeeknumMatchExcelBehaviorAndReturnsZeroForSerialDateZeroWhenWeekStartsOnSunday()
    {
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExpr("WEEKNUM(0,1)"));
        ClassicAssert.AreEqual(0, XLWorkbook.EvaluateExpr("WEEKNUM(0,17)"));
    }

    [Test]
    [Arguments]
    public void WeeknumReturnsNumberInvalidErrorOnNonSerialDates()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKNUM(-0.1)"));
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("WEEKNUM(DATE(9999,12,31)+1)")
        );
    }

    [Test]
    [Arguments(-5)]
    [Arguments(0)]
    [Arguments(3)]
    [Arguments(10)]
    [Arguments(18)]
    [Arguments(20)]
    [Arguments(22)]
    [Arguments(100)]
    public void WeeknumReturnsNumberInvalidErrorOnNonSpecifiedFlags(double flag) =>
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr($"WEEKNUM(DATE(200,1,1),{flag})")
        );

    [Test]
    public void WorkdaysMultipleHolidaysGiven()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell()
            .SetValue("Date")
            .CellBelow()
            .SetValue(new DateTime(2008, 10, 1))
            .CellBelow()
            .SetValue(151)
            .CellBelow()
            .SetValue(new DateTime(2008, 11, 26))
            .CellBelow()
            .SetValue(new DateTime(2008, 12, 4))
            .CellBelow()
            .SetValue(new DateTime(2009, 1, 21));
        XLCellValue actual = ws.Evaluate("Workday(A2,A3,A4:A6)");
        ClassicAssert.AreEqual(new DateTime(2009, 5, 5).ToSerialDateTime(), actual);
    }

    [Test]
    public void WorkdaysNoHolidaysGiven()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 151)");
        ClassicAssert.AreEqual(new DateTime(2009, 4, 30).ToSerialDateTime(), actual);

        actual = XLWorkbook.EvaluateExpr("Workday(\"2016-01-01\", -10)");
        ClassicAssert.AreEqual(new DateTime(2015, 12, 18).ToSerialDateTime(), actual);
    }

    [Test]
    public void WorkdaysOneHolidaysGiven()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            "Workday(\"10/01/2008\", 152, \"11/26/2008\")"
        );
        ClassicAssert.AreEqual(new DateTime(2009, 5, 4).ToSerialDateTime(), actual);
    }

    [Test]
    [Arguments(0, 0, 0)]
    [Arguments(0, 1, 2)]
    [Arguments(1, 1, 2)]
    [Arguments(2, 1, 3)]
    [Arguments(0, 5, 6)]
    [Arguments(2, 8, 12)]
    [Arguments(10, -1, 9)]
    [Arguments(10, -2, 6)]
    [Arguments(10, -3, 5)]
    [Arguments(9, -1, 6)]
    [Arguments(9, -2, 5)]
    [Arguments(8, -1, 6)]
    [Arguments(7, -1, 6)]
    [Arguments(6, -1, 5)]
    public void Workdays(int startDate, int dayOffset, int expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"WORKDAY({startDate}, {dayOffset})");
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Arguments(0, 1, new[] { 1 }, 2)]
    [Arguments(0, 1, new[] { 2 }, 3)]
    [Arguments(0, 5, new[] { 2, 4 }, 10)]
    [Arguments(0, 4, new[] { 2, 4, 6 }, 10)]
    [Arguments(0, 3, new[] { 2, 3, 4, 6 }, 10)]
    [Arguments(0, 2, new[] { 2, 3, 4, 5, 6 }, 10)]
    [Arguments(0, 1, new[] { 2, 3, 5 }, 4)]
    [Arguments(0, 2, new[] { 2, 3, 5 }, 6)]
    [Arguments(2, 1, new[] { 2 }, 3)]
    [Arguments(15, -1, new[] { 13 }, 12)] // 15 = Sunday
    [Arguments(100, -5, new[] { 82, 93, 94, 95, 94, 100 }, 88)]
    [Arguments(98, -2, new[] { 97 }, 95)]
    public void WorkdaysWithHoliday(int startDate, int dayOffset, int[] holidays, int expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            $"WORKDAY({startDate}, {dayOffset}, {{{string.Join(",", holidays)}}})"
        );
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Arguments("\"8/22/2008\"", 2008)]
    [Arguments("\"1/2/2006 10:45 AM\"", 2006)]
    [Arguments("0", 1900)]
    [Arguments("0.5", 1900)]
    [Arguments("1", 1900)]
    [Arguments("59", 1900)]
    [Arguments("60", 1900)]
    [Arguments("366", 1900)]
    [Arguments("367", 1901)]
    [Arguments("\"367\"", 1901)] // Test string in addition to number in TestCase before
    [Arguments("DATE(9999,12,31)+0.9", 9999)]
    [Arguments("DATE(9999,12,31)+1", XLError.NumberInvalid)]
    [Arguments("-1", XLError.NumberInvalid)]
    [Arguments("\"test\"", XLError.IncompatibleValue)]
    [Arguments("IF(TRUE,)", 1900)] // Blank
    [Arguments("TRUE", 1900)]
    [Arguments("FALSE", 1900)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("\"\"", XLError.IncompatibleValue)]
    public void Year(string value, object expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"YEAR({value})");
        ClassicAssert.AreEqual(XLCellValue.FromObject(expected), actual);
    }

    [Test]
    public void YearBlankValue()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").FormulaA1 = "YEAR(A1)";
        XLCellValue valueA2 = ws.Cell("A2").Value;
        ClassicAssert.AreEqual(1900, valueA2);
    }

    [Test]
    [Skip("Excel accepts this but XlsxSharp does not yet")]
    public void YearAcceptsMissingYearAndSubstitutesCurrentYear()
    {
        // Test providing just month and day, which should fill the year as "current year"
        double actual = XLWorkbook.EvaluateExpr("YEAR(\"8/22\")").GetNumber();
        ClassicAssert.AreEqual(DateTime.Now.Year, actual);
    }

    [Test]
    [Arguments(0, 2008, 1, 1, 2008, 3, 31, 0.25)]
    [Arguments(0, 2008, 1, 1, 2013, 3, 31, 5.25)]
    [Arguments(1, 2008, 1, 1, 2008, 3, 31, 0.24590163934426229)]
    [Arguments(1, 2008, 1, 1, 2013, 3, 31, 5.24452554744526)]
    [Arguments(1, 1900, 1, 10, 2024, 2, 29, 124.137572279657)]
    [Arguments(1, 1924, 6, 25, 2025, 2, 28, 100.67763581705)]
    [Arguments(2, 2008, 1, 1, 2008, 3, 31, 0.25)]
    [Arguments(2, 2008, 1, 1, 2013, 3, 31, 5.32222222222222)]
    [Arguments(3, 2008, 1, 1, 2008, 3, 31, 0.24657534246575341)]
    [Arguments(3, 2008, 1, 1, 2013, 3, 31, 5.24931506849315)]
    [Arguments(4, 2008, 1, 1, 2008, 3, 31, 0.24722222222222223)]
    [Arguments(4, 2008, 1, 1, 2013, 3, 31, 5.24722222222222)]
    [Arguments(0, 2006, 1, 1, 2006, 3, 26, 0.23611111111)]
    [Arguments(0, 2006, 3, 26, 2006, 1, 1, 0.23611111111)]
    [Arguments(0, 2006, 1, 1, 2006, 7, 1, 0.5)]
    [Arguments(0, 2006, 1, 1, 2007, 9, 1, 1.6666666666)]
    [Arguments(1, 2006, 1, 1, 2006, 7, 1, 0.495890411)]
    [Arguments(2, 2006, 1, 1, 2006, 7, 1, 0.5027777778)]
    [Arguments(3, 2006, 1, 1, 2006, 7, 1, 0.495890411)]
    [Arguments(4, 2006, 1, 1, 2006, 7, 1, 0.5)]
    [Arguments(1, 2004, 3, 1, 2006, 3, 1, 1.9981751825)]
    public void YearFracCalculatesFractionOfAYear(
        double basis,
        double startYear,
        double startMonth,
        double startDay,
        double endYear,
        double endMonth,
        double endDay,
        double expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            (double)
                XLWorkbook.EvaluateExpr(
                    $"YEARFRAC(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),{basis})"
                ),
            XLHelper.Epsilon
        );
    }

    [Test]
    public void YearFracDatesMustFitInDateSystemRange()
    {
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("YEARFRAC(-0.1,10)"));
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("YEARFRAC(0,-0.1)"));
    }

    [Test]
    public void YearFracBasisMustBeBetween0And4()
    {
        ClassicAssert.AreEqual(
            XLError.NumberInvalid,
            XLWorkbook.EvaluateExpr("YEARFRAC(0,10,-0.1)")
        );
        ClassicAssert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("YEARFRAC(0,10,5)"));
    }
}
