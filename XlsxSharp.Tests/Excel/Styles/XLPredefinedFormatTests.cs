using System.Collections.Generic;
using System.Globalization;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

internal class XlPredefinedFormatTests
{
    [Test]
    [MethodDataSource(nameof(FormattedStringTestCases))]
    [Arguments(
        0.25,
        (int)XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
        "cs-CZ",
        "6:00 dop.",
        Skip = "ExcelNumberFormat always uses invariant culture for AM/PM."
    )]
    [Arguments(
        0.75,
        (int)XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
        "cs-CZ",
        "6:00 odp.",
        Skip = "ExcelNumberFormat always uses invariant culture for AM/PM."
    )]
    [Arguments(
        0.8,
        (int)XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
        "cs-CZ",
        "7:12:00 odp.",
        Skip = "ExcelNumberFormat always uses invariant culture for AM/PM."
    )]
    public void Predefined_formats_are_correctly_formatted(
        double value,
        int predefinedFormatId,
        string cultureName,
        string expectedText
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").SetValue(value).Style.NumberFormat.SetNumberFormatId(predefinedFormatId);

        string formattedText = ws.Cell("A1")
            .GetFormattedString(CultureInfo.GetCultureInfo(cultureName));

        ClassicAssert.AreEqual(expectedText, formattedText);
    }

    public static IEnumerable<(double, int, string, string)> FormattedStringTestCases()
    {
        const string en = "en-US";
        const string invariant = "";

        yield return (14.25, (int)XLPredefinedFormat.DateTime.Hour12MinutesAmPm, en, "6:00 AM");
        yield return (0.5, (int)XLPredefinedFormat.DateTime.Hour12MinutesAmPm, en, "12:00 PM");
        yield return (14.75, (int)XLPredefinedFormat.DateTime.Hour12MinutesAmPm, en, "6:00 PM");
        yield return (
            14.25,
            (int)XLPredefinedFormat.DateTime.Hour12MinutesAmPm,
            invariant,
            "6:00 AM"
        );

        yield return (
            7.123,
            (int)XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
            en,
            "2:57:07 AM"
        );
        yield return (
            2.5,
            (int)XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
            en,
            "12:00:00 PM"
        );
        yield return (
            2.99,
            (int)XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
            en,
            "11:45:36 PM"
        );
        yield return (
            0.75,
            (int)XLPredefinedFormat.DateTime.Hour12MinutesSecondsAmPm,
            invariant,
            "6:00:00 PM"
        );
    }
}
