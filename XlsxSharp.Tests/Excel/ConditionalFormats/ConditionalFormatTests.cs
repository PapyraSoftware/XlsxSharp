using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;

namespace XlsxSharp.Tests.Excel.ConditionalFormats;

public class ConditionalFormatTests
{
    [Test]
    public void MaintainConditionalFormattingOrder() =>
        // The input file contains duplicates of same dxf and thus the output also contains some duplicate dxf.
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\ConditionalFormattingOrder\inputfile.xlsx",
            @"Other\StyleReferenceFiles\ConditionalFormattingOrder\ConditionalFormattingOrder.xlsx"
        );

    [Test]
    [Arguments(true, 7)]
    [Arguments(false, 8)]
    public void SaveOptionAffectsConsolidationConditionalFormatRanges(
        bool consolidateConditionalFormatRanges,
        int expectedCount
    )
    {
        SaveOptions options = new()
        {
            ConsolidateConditionalFormatRanges = consolidateConditionalFormatRanges,
        };

        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");

        ws.Range("D2:D3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("B2:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("E2:E6")
            .AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.Red)
            .HighestValue(XLColor.Blue);
        ws.Range("F2:F6")
            .AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.Red)
            .HighestValue(XLColor.Blue);
        ws.Range("G2:G7")
            .AddConditionalFormat()
            .WhenIsUnique()
            .Fill.SetBackgroundColor(XLColor.Blue);
        ws.Range("H2:H7")
            .AddConditionalFormat()
            .WhenIsUnique()
            .Fill.SetBackgroundColor(XLColor.Blue);
        ws.Range("I2:I6").AddConditionalFormat().WhenContains("test");
        ws.Range("J2:J6").AddConditionalFormat().WhenContains("test");
        using (MemoryStream ms = new())
        {
            wb.SaveAs(ms, options);
            XLWorkbook wb_saved = new(ms);
            ClassicAssert.AreEqual(
                expectedCount,
                wb_saved.Worksheet("Sheet").ConditionalFormats.Count()
            );
        }
    }

    [Test]
    [Arguments(true, 1)]
    [Arguments(false, 2)]
    public void SaveOptionAffectsConsolidationDataValidationRanges(
        bool consolidateDataValidationRanges,
        int expectedCount
    )
    {
        SaveOptions options = new()
        {
            ConsolidateDataValidationRanges = consolidateDataValidationRanges,
        };

        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        ws.Range("C2:C5").CreateDataValidation().Decimal.Between(1, 5);
        ws.Range("D2:D5").CreateDataValidation().Decimal.Between(1, 5);

        using (MemoryStream ms = new())
        {
            wb.SaveAs(ms, options);
            XLWorkbook wb_saved = new(ms);
            ClassicAssert.AreEqual(
                expectedCount,
                wb_saved.Worksheet("Sheet").DataValidations.Count()
            );
        }
    }

    [Test]
    [Arguments("en-US")]
    [Arguments("fr-FR")]
    [Arguments("ru-RU")]
    public void SaveConditionalFormatCultureIndependent(string culture)
    {
        using (MemoryStream ms = new())
        {
            double expectedValue = 1.5;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.AddWorksheet();
                int i = 1;
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenEquals(expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenNotEquals(expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenGreaterThan(expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenLessThan(expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenEqualOrGreaterThan(expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenEqualOrLessThan(expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenBetween(expectedValue, expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell(i++, 1)
                    .AddConditionalFormat()
                    .WhenNotBetween(expectedValue, expectedValue)
                    .Fill.SetBackgroundColor(XLColor.Red);

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IEnumerable<string> conditionalFormatValues = ws
                    .ConditionalFormats.SelectMany(cf => cf.Values.Values)
                    .Select(v => v.Value)
                    .Distinct();

                ClassicAssert.AreEqual(1, conditionalFormatValues.Count());
                ClassicAssert.AreEqual("1.5", conditionalFormatValues.Single());
            }
        }
    }

    [Test]
    public void CellIsTypeReadsOnlyRequiredFormulaArguments()
    {
        // The CellIs uses formula tags as arguments. Some producers generate extra empty
        // formula tags and ClosedXml should be able to load CellIs conditional formatting
        // with such extra tags without an exception. The test file has been modified to
        // include extra formula tags and test checks that extra tags are ignored.
        TestHelper.LoadAndAssert(
            (_, ws) =>
            {
                AssertFormulaArgs(ws, XLCFOperator.Between, "$D$2", "$E$2");
                AssertFormulaArgs(ws, XLCFOperator.NotBetween, "$D$3", "$E$3");
                AssertFormulaArgs(ws, XLCFOperator.GreaterThan, "$D$4");
                AssertFormulaArgs(ws, XLCFOperator.LessThan, "$D$5");
                AssertFormulaArgs(ws, XLCFOperator.Equal, "$D$6");
            },
            @"Other\ConditionalFormats\Extra_formulas_CellIs_type.xlsx"
        );

        static void AssertFormulaArgs(
            IXLWorksheet ws,
            XLCFOperator cfOperator,
            params string[] expectedFormulas
        )
        {
            IXLConditionalFormat cf = ws.ConditionalFormats.Single(cf =>
                cf.ConditionalFormatType == XLConditionalFormatType.CellIs
                && cf.Operator == cfOperator
            );
            ClassicAssert.AreEqual(expectedFormulas.Length, cf.Values.Count);
            CollectionAssert.AreEqual(expectedFormulas, cf.Values.Select(v => v.Value.Value));
        }
    }

    [Test]
    public void ExpressionTypeSkipsEmptyFormulaTags()
    {
        // The Expression uses formula tag as arguments. Some producers generate extra empty
        // formula tags and ClosedXml should be able to load Expression conditional formatting
        // with such extra tags without an exception. The test file has been modified to
        // include extra formula tags and test checks that extra tags are ignored.
        TestHelper.LoadAndAssert(
            (_, ws) =>
            {
                AssertFormulaArgs(ws, "A1:A1", "$C$1=5");
                AssertFormulaArgs(ws, "A2:A2", "$C$2=4");
            },
            @"Other\ConditionalFormats\Extra_formulas_Expression_type.xlsx"
        );

        static void AssertFormulaArgs(IXLWorksheet ws, string range, string expectedFormula)
        {
            IXLConditionalFormat cf = ws.ConditionalFormats.Single(cf =>
                cf.ConditionalFormatType == XLConditionalFormatType.Expression
                && cf.Range.RangeAddress.ToString() == range
            );
            ClassicAssert.AreEqual(1, cf.Values.Count);
            CollectionAssert.AreEqual(expectedFormula, cf.Values[1].Value);
        }
    }

    [Test]
    public void RangeSetterThrowsOnRangeFromDifferentWorksheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        IXLConditionalFormat cf = ws1.Range("A1").AddConditionalFormat();
        IXLWorksheet ws2 = wb.AddWorksheet();
        IXLRange differentSheetRange = ws2.Range("B5");

        ClassicAssert.Throws<ArgumentException>(() => cf.Range = differentSheetRange);
    }

    [Test]
    public void RangesSetterThrowsOnRangeFromDifferentWorksheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        IXLConditionalFormat cf = ws1.Range("A1").AddConditionalFormat();

        IXLRanges ranges = ws1.Ranges("B1");
        IXLRange differentSheetRange = wb.AddWorksheet().Range("C1");
        ranges.Add(differentSheetRange);

        ArgumentException ex = ClassicAssert.Throws<ArgumentException>(() => cf.Ranges = ranges);
        StringAssert.Contains("must be from worksheet", ex.Message);
    }

    [Test]
    public void RangesSetterThrowsOnEmptyRanges()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        IXLConditionalFormat cf = ws1.Range("A1").AddConditionalFormat();

        IXLRanges emptyRanges = ws1.Ranges("");

        ArgumentException ex = ClassicAssert.Throws<ArgumentException>(() =>
            cf.Ranges = emptyRanges
        );
        StringAssert.Contains("empty", ex.Message);
    }

    [Test]
    public void TextFormatsAreLoadedAsValue()
    {
        // Issue #2690: Ensure that CF for text are loaded as a value and not as a formula.
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.AddConditionalFormat()
                    .WhenContains("Contains")
                    .Fill.SetBackgroundColor(XLColor.Red);

                ws.AddConditionalFormat()
                    .WhenNotContains("Not Contains")
                    .Fill.SetBackgroundColor(XLColor.Green);

                ws.AddConditionalFormat()
                    .WhenStartsWith("Starts With")
                    .Fill.SetBackgroundColor(XLColor.Blue);

                ws.AddConditionalFormat()
                    .WhenEndsWith("Ends With")
                    .Fill.SetBackgroundColor(XLColor.Black);
            },
            (_, ws) =>
            {
                AssertTextCfValue(ws, XLConditionalFormatType.ContainsText, "Contains");
                AssertTextCfValue(ws, XLConditionalFormatType.NotContainsText, "Not Contains");
                AssertTextCfValue(ws, XLConditionalFormatType.StartsWith, "Starts With");
                AssertTextCfValue(ws, XLConditionalFormatType.EndsWith, "Ends With");
            }
        );
        return;

        static void AssertTextCfValue(IXLWorksheet ws, XLConditionalFormatType type, string text)
        {
            IXLConditionalFormat cf = ws.ConditionalFormats.Single(x =>
                x.ConditionalFormatType == type
            );
            ClassicAssert.AreEqual(1, cf.Values.Count);
            ClassicAssert.AreEqual(text, cf.Values[1].Value);
            ClassicAssert.IsFalse(cf.Values[1].IsFormula);
        }
    }
}
