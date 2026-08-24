using System.Globalization;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.AutoFilters;

public class AutoFilterTests
{
    [Test]
    public void AutoFilterExpandsWithTable()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.FirstCell()
                .SetValue("Categories")
                .CellBelow()
                .SetValue("1")
                .CellBelow()
                .SetValue("2");

            IXLTable table = ws.RangeUsed().CreateTable();

            List<int> listOfArr = [3, 4, 5, 6];

            table.DataRange.InsertRowsBelow(listOfArr.Count - table.DataRange.RowCount());
            table.DataRange.FirstCell().InsertData(listOfArr);

            ClassicAssert.AreEqual("A1:A5", table.AutoFilter.Range.RangeAddress.ToStringRelative());
            ClassicAssert.AreEqual(5, table.AutoFilter.VisibleRows.Count());
        }
    }

    [Test]
    public void AutoFilterSortWhenNotInFirstRow()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Cell(3, 3)
                .SetValue("Names")
                .CellBelow()
                .SetValue("Manuel")
                .CellBelow()
                .SetValue("Carlos")
                .CellBelow()
                .SetValue("Dominic");
            ws.RangeUsed().SetAutoFilter().Sort();
            ClassicAssert.AreEqual("Carlos", ws.Cell(4, 3).GetText());
        }
    }

    [Test]
    public void CanClearAutoFilter()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("AutoFilter");
        ws.Cell("A1").Value = "Names";
        ws.Cell("A2").Value = "John";
        ws.Cell("A3").Value = "Hank";
        ws.Cell("A4").Value = "Dagny";

        ws.AutoFilter.Clear(); // We should be able to clear a filter even if it hasn't been set.
        ClassicAssert.IsTrue(!ws.AutoFilter.IsEnabled);

        ws.RangeUsed().SetAutoFilter();
        ClassicAssert.IsTrue(ws.AutoFilter.IsEnabled);

        ws.AutoFilter.Clear();
        ClassicAssert.IsTrue(!ws.AutoFilter.IsEnabled);
    }

    [Test]
    public void CanClearAutoFilter2()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Names";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.SetAutoFilter(false);
            ClassicAssert.IsTrue(!ws.AutoFilter.IsEnabled);

            ws.RangeUsed().SetAutoFilter();
            ClassicAssert.IsTrue(ws.AutoFilter.IsEnabled);

            ws.RangeUsed().SetAutoFilter(false);
            ClassicAssert.IsTrue(!ws.AutoFilter.IsEnabled);
        }
    }

    [Test]
    public void CanCopyAutoFilterToNewSheetOnNewWorkbook()
    {
        using (MemoryStream ms1 = new())
        using (MemoryStream ms2 = new())
        {
            using (XLWorkbook wb1 = new())
            using (XLWorkbook wb2 = new())
            {
                IXLWorksheet ws = wb1.Worksheets.Add("AutoFilter");
                ws.Cell("A1").Value = "Names";
                ws.Cell("A2").Value = "John";
                ws.Cell("A3").Value = "Hank";
                ws.Cell("A4").Value = "Dagny";

                ws.RangeUsed().SetAutoFilter();

                wb1.SaveAs(ms1);

                ws.CopyTo(wb2, ws.Name);
                wb2.SaveAs(ms2);
            }

            using (XLWorkbook wb2 = new(ms2))
            {
                ClassicAssert.IsTrue(wb2.Worksheets.First().AutoFilter.IsEnabled);
            }
        }
    }

    [Test]
    public void CannotAddAutoFilterOverExistingTable()
    {
        using XLWorkbook wb = new();

        var data = Enumerable.Range(1, 10).Select(i => new { Index = i, String = $"String {i}" });

        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell().InsertTable(data);

        ClassicAssert.Throws<InvalidOperationException>(() => ws.RangeUsed().SetAutoFilter());
    }

    [Test]
    [Arguments("A1:A4")]
    [Arguments("A1:B4")]
    [Arguments("A1:C4")]
    public void AutoFilterRangeRemainsValidOnInsertColumn(string rangeAddress)
    {
        //Arrange
        using (MemoryStream ms1 = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add("AutoFilter");
                ws.Cell("A1").Value = "Ids";
                ws.Cell("B1").Value = "Names";
                ws.Cell("B2").Value = "John";
                ws.Cell("B3").Value = "Hank";
                ws.Cell("B4").Value = "Dagny";
                ws.Cell("C1").Value = "Phones";

                ws.Range("B1:B4").SetAutoFilter(true);

                //Act
                IXLRange range = ws.Range(rangeAddress);
                range.InsertColumnsBefore(1);

                //Assert
                ClassicAssert.IsTrue(ws.AutoFilter.Range.RangeAddress.IsValid);
            }
        }
    }

    [Test]
    public void AutoFilterVisibleRows()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Cell(3, 3)
                .SetValue("Names")
                .CellBelow()
                .SetValue("Manuel")
                .CellBelow()
                .SetValue("Carlos")
                .CellBelow()
                .SetValue("Dominic");

            IXLAutoFilter autoFilter = ws.RangeUsed().SetAutoFilter();

            autoFilter.Column(1).AddFilter("Carlos");

            ClassicAssert.AreEqual("Carlos", ws.Cell(5, 3).GetText());
            ClassicAssert.AreEqual(2, autoFilter.VisibleRows.Count());
            ClassicAssert.AreEqual(3, autoFilter.VisibleRows.First().WorksheetRow().RowNumber());
            ClassicAssert.AreEqual(5, autoFilter.VisibleRows.Last().WorksheetRow().RowNumber());
        }
    }

    [Test]
    public void ReapplyAutoFilter()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Cell(3, 3)
                .SetValue("Names")
                .CellBelow()
                .SetValue("Manuel")
                .CellBelow()
                .SetValue("Carlos")
                .CellBelow()
                .SetValue("Dominic")
                .CellBelow()
                .SetValue("Jose");

            IXLAutoFilter autoFilter = ws.RangeUsed().SetAutoFilter();

            autoFilter.Column(1).AddFilter("Carlos");

            ClassicAssert.AreEqual(3, autoFilter.HiddenRows.Count());

            // Unhide the rows so that the table is out of sync with the filter
            autoFilter.HiddenRows.ForEach(r => r.WorksheetRow().Unhide());
            ClassicAssert.False(autoFilter.HiddenRows.Any());

            autoFilter.Reapply();
            ClassicAssert.AreEqual(3, autoFilter.HiddenRows.Count());
        }
    }

    [Test]
    public void CanLoadAutoFilterWithThousandsSeparator()
    {
        CultureInfo backupCulture = Thread.CurrentThread.CurrentCulture;

        try
        {
            // Set thread culture to French, which should format numbers using a space as thousands separator
            CultureInfo culture = CultureInfo.CreateSpecificCulture("fr-FR");

            // The value in sheet that will be compared with autofilter value is a number
            // `10000`. That number will be formatted using culture to `10 000.00` thanks to
            // modified properties of culture - period instead of a comma for decimal separator
            // and space as group separator. The formatted number will thus match with the
            // filter value.
            culture.NumberFormat.NumberDecimalSeparator = ".";
            culture.NumberFormat.NumberGroupSeparator = " ";

            Thread.CurrentThread.CurrentCulture = culture;

            using (
                Stream stream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(
                        @"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx"
                    )
                )
            )
            using (XLWorkbook wb = new(stream))
            {
                IXLWorksheet ws = wb.Worksheets.First();

                // Regular filter compares values as strings, doesn't convert to XLCellValue,
                // so the value is read from the file as a text despite looking like a number.
                ClassicAssert.AreEqual(
                    "10 000.00",
                    ((XLAutoFilter)ws.AutoFilter).Column(1).Single().Value
                );
                ClassicAssert.AreEqual(2, ws.AutoFilter.VisibleRows.Count());

                ws.AutoFilter.Reapply();
                ClassicAssert.AreEqual(2, ws.AutoFilter.VisibleRows.Count());
            }

            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

            using (
                Stream stream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(
                        @"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx"
                    )
                )
            )
            using (XLWorkbook wb = new(stream))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                ClassicAssert.AreEqual(
                    "10 000.00",
                    ((XLAutoFilter)ws.AutoFilter).Column(1).Single().Value
                );

                List<XLCellValue> v =
                [
                    .. ws.AutoFilter.VisibleRows.Select(r => r.FirstCell().Value),
                ];
                ClassicAssert.AreEqual(2, ws.AutoFilter.VisibleRows.Count());

                ws.AutoFilter.Reapply();
                ClassicAssert.AreEqual(1, ws.AutoFilter.VisibleRows.Count());
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = backupCulture;
        }
    }

    [Test]
    public void Issue1917NotContainsFilter()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Test");
                ws.Cell(1, 1).SetValue("StringCol");

                for (int i = 0; i < 5; i++)
                {
                    ws.Cell(i + 2, 1).SetValue($"String{i}");
                }

                IXLAutoFilter autoFilter = ws.RangeUsed().SetAutoFilter();

                autoFilter.Column(1).NotContains("String3");
                ClassicAssert.AreEqual(1, autoFilter.HiddenRows.Count());

                wb.SaveAs(ms);
            }

            ms.Position = 0;
            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.Worksheet("Test");
                IXLAutoFilter autoFilter = ws.AutoFilter;

                autoFilter.Reapply();
                ClassicAssert.AreEqual(1, autoFilter.HiddenRows.Count());
            }
        }
    }

    [Test]
    [Arguments("ends")]
    [Arguments("begins")]
    [Arguments("equal")]
    [Arguments("contains")]
    public void NotStringFilter(string type)
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Test");
                ws.Cell(1, 1).SetValue("StringCol");

                for (int i = 0; i < 5; i++)
                {
                    ws.Cell(i + 2, 1).SetValue($"{i}-String{i}");
                }

                ws.Columns().AdjustToContents();
                IXLAutoFilter autoFilter = ws.RangeUsed().SetAutoFilter();

                switch (type)
                {
                    case "ends":
                        autoFilter.Column(1).NotEndsWith("3");
                        break;
                    case "begins":
                        autoFilter.Column(1).NotBeginsWith("3");
                        break;
                    case "equal":
                        autoFilter.Column(1).NotEqualTo("3-String3");
                        break;
                    case "contains":
                        autoFilter.Column(1).NotContains("3-");
                        break;
                }
                ClassicAssert.AreEqual(1, autoFilter.HiddenRows.Count());

                wb.SaveAs(ms);
            }

            ms.Position = 0;
            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.Worksheet("Test");
                IXLAutoFilter autoFilter = ws.AutoFilter;

                autoFilter.Reapply();
                ClassicAssert.AreEqual(1, autoFilter.HiddenRows.Count());
            }
        }
    }
}
