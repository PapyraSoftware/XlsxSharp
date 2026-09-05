using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.IO;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Tables;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XlsxSharp.Tests.Excel.IO;

/// <summary>
/// Checks the worksheet enumerations against the SDK conversions they replace, in both
/// directions: every string the SDK serialises has to parse to the value its converter produced,
/// and every workbook model value has to write the string the SDK wrote for it. A misspelling or
/// a swapped pair fails here rather than turning into a worksheet that loads with the wrong page
/// order or saves with the wrong icon set.
/// </summary>
public class WorksheetXmlEnumsTests
{
    [Test]
    public void PageOrientationMatchesTheSdk() =>
        AssertParity(
            (OrientationValues v) => v.ToXlsxSharp(),
            (XLPageOrientation v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParsePageOrientation,
            v => v.ToXml()
        );

    [Test]
    public void PageOrderMatchesTheSdk() =>
        AssertParity(
            (PageOrderValues v) => v.ToXlsxSharp(),
            (XLPageOrderValues v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParsePageOrder,
            v => v.ToXml()
        );

    [Test]
    public void ShowCommentsMatchesTheSdk() =>
        AssertParity(
            (CellCommentsValues v) => v.ToXlsxSharp(),
            (XLShowCommentsValues v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseShowComments,
            v => v.ToXml()
        );

    [Test]
    public void PrintErrorMatchesTheSdk() =>
        AssertParity(
            (PrintErrorValues v) => v.ToXlsxSharp(),
            (XLPrintErrorValues v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParsePrintError,
            v => v.ToXml()
        );

    [Test]
    public void SheetViewTypeMatchesTheSdk() =>
        AssertParity(
            (SheetViewValues v) => v.ToXlsxSharp(),
            (XLSheetViewOptions v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseSheetViewType,
            v => v.ToXml()
        );

    [Test]
    public void AllowedValuesMatchesTheSdk() =>
        AssertParity(
            (DataValidationValues v) => v.ToXlsxSharp(),
            (XLAllowedValues v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseAllowedValues,
            v => v.ToXml()
        );

    [Test]
    public void ErrorStyleMatchesTheSdk() =>
        AssertParity(
            (DataValidationErrorStyleValues v) => v.ToXlsxSharp(),
            (XLErrorStyle v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseErrorStyle,
            v => v.ToXml()
        );

    [Test]
    public void DataValidationOperatorMatchesTheSdk() =>
        AssertParity(
            (DataValidationOperatorValues v) => v.ToXlsxSharp(),
            (XLOperator v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseDataValidationOperator,
            v => v.ToXml()
        );

    [Test]
    public void ConditionalFormatTypeMatchesTheSdk() =>
        AssertParity(
            (ConditionalFormatValues v) => v.ToXlsxSharp(),
            (XLConditionalFormatType v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseConditionalFormatType,
            v => v.ToXml()
        );

    [Test]
    public void CfOperatorMatchesTheSdk() =>
        AssertParity(
            (ConditionalFormattingOperatorValues v) => v.ToXlsxSharp(),
            (XLCFOperator v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseCfOperator,
            v => v.ToXml()
        );

    [Test]
    public void CfContentTypeMatchesTheSdk() =>
        AssertParity(
            (ConditionalFormatValueObjectValues v) => v.ToXlsxSharp(),
            (XLCFContentType v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseCfContentType,
            v => v.ToXml()
        );

    [Test]
    public void IconSetStyleMatchesTheSdk() =>
        AssertParity(
            (IconSetValues v) => v.ToXlsxSharp(),
            (XLIconSetStyle v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseIconSetStyle,
            v => v.ToXml()
        );

    [Test]
    public void TimePeriodMatchesTheSdk() =>
        AssertParity(
            (TimePeriodValues v) => v.ToXlsxSharp(),
            (XLTimePeriod v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseTimePeriod,
            v => v.ToXml()
        );

    [Test]
    public void FilterOperatorMatchesTheSdk() =>
        AssertParity(
            (FilterOperatorValues v) => v.ToXlsxSharp(),
            (XLFilterOperator v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseFilterOperator,
            v => v.ToXml()
        );

    [Test]
    public void FilterDynamicTypeMatchesTheSdk() =>
        AssertParity(
            (DynamicFilterValues v) => v.ToXlsxSharp(),
            (XLFilterDynamicType v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseFilterDynamicType,
            v => v.ToXml()
        );

    [Test]
    public void DateTimeGroupingMatchesTheSdk() =>
        AssertParity(
            (DateTimeGroupingValues v) => v.ToXlsxSharp(),
            (XLDateTimeGrouping v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseDateTimeGrouping,
            v => v.ToXml()
        );

    [Test]
    public void TotalsRowFunctionMatchesTheSdk() =>
        AssertParity(
            (TotalsRowFunctionValues v) => v.ToXlsxSharp(),
            (XLTotalsRowFunction v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseTotalsRowFunction,
            v => v.ToXml()
        );

    [Test]
    public void PhoneticTypeMatchesTheSdk() =>
        AssertParity(
            (PhoneticValues v) => v.ToXlsxSharp(),
            (XLPhoneticType v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParsePhoneticType,
            v => v.ToXml()
        );

    [Test]
    public void PhoneticAlignmentMatchesTheSdk()
    {
        AssertParsesEveryValue(
            (PhoneticAlignmentValues v) => v.ToXlsxSharp(),
            WorksheetXmlEnums.ParsePhoneticAlignment
        );

        // The SDK had no XLPhoneticAlignment -> enum conversion, only the string the writer used.
        foreach (XLPhoneticAlignment value in Enum.GetValues<XLPhoneticAlignment>())
        {
            ClassicAssert.AreEqual(value.ToOpenXmlString(), value.ToXml(), value.ToString());
        }
    }

    [Test]
    public void SparklineTypeMatchesTheSdk() =>
        AssertParity(
            (X14.SparklineTypeValues v) => v.ToXlsxSharp(),
            (XLSparklineType v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseSparklineType,
            v => v.ToXml()
        );

    [Test]
    public void SparklineAxisMinMaxMatchesTheSdk() =>
        AssertParity(
            (X14.SparklineAxisMinMaxValues v) => v.ToXlsxSharp(),
            (XLSparklineAxisMinMax v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseSparklineAxisMinMax,
            v => v.ToXml()
        );

    [Test]
    public void DisplayBlanksAsMatchesTheSdk() =>
        AssertParity(
            (X14.DisplayBlanksAsValues v) => v.ToXlsxSharp(),
            (XLDisplayBlanksAsValues v) => v.ToOpenXml(),
            WorksheetXmlEnums.ParseDisplayBlanksAs,
            v => v.ToXml()
        );

    [Test]
    public void PicturePlacementMatchesTheSdk() =>
        AssertWritesEveryValue((XLPicturePlacement v) => v.ToOpenXml(), v => v.ToXml());

    [Test]
    public void AnUnknownValueIsRejected() =>
        ClassicAssert.Throws<XlsxSharp.IO.PartStructureException>(() =>
            WorksheetXmlEnums.ParsePageOrder("sidewaysThenUp")
        );

    private static void AssertParity<TSdk, TXl>(
        Func<TSdk, TXl> convertWithSdk,
        Func<TXl, TSdk> writeWithSdk,
        Func<string, TXl> parse,
        Func<TXl, string> toXml
    )
        where TSdk : struct, IEnumValue
        where TXl : struct, Enum
    {
        AssertParsesEveryValue(convertWithSdk, parse);
        AssertWritesEveryValue(writeWithSdk, toXml);
    }

    /// <summary>
    /// Every string the SDK serialises parses to the value the SDK conversion produced. A value
    /// the old conversion had no workbook model equivalent for - the schema has three dozen
    /// dynamic filters and the model knows two - has to stay unreadable rather than become a
    /// silently wrong one.
    /// </summary>
    private static void AssertParsesEveryValue<TSdk, TXl>(
        Func<TSdk, TXl> convertWithSdk,
        Func<string, TXl> parse
    )
        where TSdk : struct, IEnumValue
    {
        int checkedValues = 0;

        foreach ((string name, TSdk value) in ValuesOf<TSdk>())
        {
            string xml = ((IEnumValue)value).Value;
            TXl expected;
            try
            {
                expected = convertWithSdk(value);
            }
            catch (KeyNotFoundException)
            {
                ClassicAssert.Throws<XlsxSharp.IO.PartStructureException>(
                    () => parse(xml),
                    $"{typeof(TSdk).Name}.{name} ('{xml}')"
                );
                checkedValues++;
                continue;
            }

            ClassicAssert.AreEqual(expected!, parse(xml)!, $"{typeof(TSdk).Name}.{name} ('{xml}')");
            checkedValues++;
        }

        ClassicAssert.IsTrue(checkedValues > 0, $"no values found on {typeof(TSdk).Name}");
    }

    /// <summary>
    /// Every workbook model value writes the string the SDK wrote for it.
    /// </summary>
    private static void AssertWritesEveryValue<TXl, TSdk>(
        Func<TXl, TSdk> writeWithSdk,
        Func<TXl, string> toXml
    )
        where TSdk : struct, IEnumValue
        where TXl : struct, Enum
    {
        foreach (TXl value in Enum.GetValues<TXl>())
        {
            ClassicAssert.AreEqual(
                ((IEnumValue)writeWithSdk(value)).Value,
                toXml(value),
                $"{typeof(TXl).Name}.{value}"
            );
        }
    }

    private static IEnumerable<(string Name, TSdk Value)> ValuesOf<TSdk>()
        where TSdk : struct, IEnumValue
    {
        foreach (
            PropertyInfo property in typeof(TSdk).GetProperties(
                BindingFlags.Public | BindingFlags.Static
            )
        )
        {
            if (property.PropertyType == typeof(TSdk))
            {
                yield return (property.Name, (TSdk)property.GetValue(null)!);
            }
        }
    }
}
