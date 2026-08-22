using System.Collections;
using System.Globalization;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

internal class XlPredefinedFormatTests
{
    [TestCaseSource(nameof(FormattedStringTestCases))]
    public void Predefined_formats_are_correctly_formatted(
        double value,
        int predefinedFormatId,
        CultureInfo culture,
        string expectedText
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").SetValue(value).Style.NumberFormat.SetNumberFormatId(predefinedFormatId);

        string formattedText = ws.Cell("A1").GetFormattedString(culture);

        Assert.AreEqual(expectedText, formattedText);
    }

    public static IEnumerable FormattedStringTestCases
    {
        get
        {
            CultureInfo en = CultureInfo.GetCultureInfo("en-US");
            CultureInfo cs = CultureInfo.GetCultureInfo("cs-CZ");
            CultureInfo invariant = CultureInfo.InvariantCulture;

            yield return new TestCaseData(
                14.25,
                XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
                en,
                "6:00 AM"
            );
            yield return new TestCaseData(
                0.5,
                XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
                en,
                "12:00 PM"
            );
            yield return new TestCaseData(
                14.75,
                XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
                en,
                "6:00 PM"
            );
            yield return new TestCaseData(
                0.25,
                XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
                cs,
                "6:00 dop."
            ).Ignore("ExcelNumberFormat always uses invariant culture for AM/PM.");
            yield return new TestCaseData(
                0.75,
                XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
                cs,
                "6:00 odp."
            ).Ignore("ExcelNumberFormat always uses invariant culture for AM/PM.");
            yield return new TestCaseData(
                14.25,
                XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
                invariant,
                "6:00 AM"
            );

            yield return new TestCaseData(
                7.123,
                XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
                en,
                "2:57:07 AM"
            );
            yield return new TestCaseData(
                2.5,
                XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
                en,
                "12:00:00 PM"
            );
            yield return new TestCaseData(
                2.99,
                XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
                en,
                "11:45:36 PM"
            );
            yield return new TestCaseData(
                0.8,
                XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
                cs,
                "7:12:00 odp."
            ).Ignore("ExcelNumberFormat always uses invariant culture for AM/PM.");
            yield return new TestCaseData(
                0.75,
                XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
                invariant,
                "6:00:00 PM"
            );
        }
    }
}
