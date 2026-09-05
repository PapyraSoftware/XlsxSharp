using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel;
using XlsxSharp.Excel.IO;

namespace XlsxSharp.Tests.Excel.PivotTables;

/// <summary>
/// Checks the pivot enumerations against the SDK conversions they replace. For every value the
/// SDK knows, the string it serialises has to parse to the same workbook model value the SDK
/// converter produced - so a misspelled string or a swapped pair fails here rather than turning
/// into a pivot table that loads with the wrong aggregation.
/// </summary>
public class PivotXmlEnumsTests
{
    [Test]
    public void SubtotalMatchesTheSdk() =>
        AssertParity(
            (DataConsolidateFunctionValues v) => v.ToXlsxSharp(),
            PivotXmlEnums.ParseSubtotal
        );

    [Test]
    public void ShowDataAsMatchesTheSdk() =>
        AssertParity((ShowDataAsValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParseShowDataAs);

    [Test]
    public void PivotAreaTypeMatchesTheSdk() =>
        AssertParity((PivotAreaValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParsePivotAreaType);

    [Test]
    public void ItemTypeMatchesTheSdk() =>
        AssertParity((ItemValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParseItemType);

    [Test]
    public void FieldSortMatchesTheSdk() =>
        AssertParity((FieldSortValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParseFieldSort);

    [Test]
    public void FormatActionMatchesTheSdk() =>
        AssertParity((FormatActionValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParseFormatAction);

    [Test]
    public void CfScopeMatchesTheSdk() =>
        AssertParity((ScopeValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParseCfScope);

    [Test]
    public void CfRuleTypeMatchesTheSdk() =>
        AssertParity((RuleValues v) => v.ToXlsxSharp(), PivotXmlEnums.ParseCfRuleType);

    [Test]
    public void AnUnknownValueIsRejected() =>
        ClassicAssert.Throws<XlsxSharp.IO.PartStructureException>(() =>
            PivotXmlEnums.ParseSubtotal("notAFunction")
        );

    private static void AssertParity<TSdk, TXl>(
        Func<TSdk, TXl> convertWithSdk,
        Func<string, TXl> parse
    )
        where TSdk : struct, IEnumValue
    {
        int checkedValues = 0;

        foreach (
            PropertyInfo property in typeof(TSdk).GetProperties(
                BindingFlags.Public | BindingFlags.Static
            )
        )
        {
            if (property.PropertyType != typeof(TSdk))
            {
                continue;
            }

            TSdk value = (TSdk)property.GetValue(null)!;
            string xml = ((IEnumValue)value).Value;

            ClassicAssert.AreEqual(
                convertWithSdk(value)!,
                parse(xml)!,
                $"{typeof(TSdk).Name}.{property.Name} ('{xml}')"
            );

            checkedValues++;
        }

        // A conversion that silently checked nothing would pass every assertion above.
        ClassicAssert.IsTrue(checkedValues > 0, $"no values found on {typeof(TSdk).Name}");
    }
}
