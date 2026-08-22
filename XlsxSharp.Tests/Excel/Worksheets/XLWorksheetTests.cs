using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using NUnit.Framework;
using XlsxSharp.Examples;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.Worksheets;

[TestFixture]
public class XLWorksheetTests
{
    private static readonly char[] illegalWorksheetCharacters =
        "\u0000\u0003:\\/?*[]".ToCharArray();

    [Test]
    public void ColumnCountTime()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        DateTime start = DateTime.Now;
        ws.ColumnCount();
        DateTime end = DateTime.Now;
        Assert.IsTrue((end - start).TotalMilliseconds < 500);
    }

    [Test]
    public void CopyConditionalFormatsCount()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:C3")
            .AddConditionalFormat()
            .WhenContains("1")
            .Fill.SetBackgroundColor(XLColor.Blue);
        ws.Range("A1:C3").Value = 1;
        IXLWorksheet ws2 = ws.CopyTo("Sheet2");
        Assert.AreEqual(1, ws2.ConditionalFormats.Count());
    }

    [Test]
    public void CopyColumnVisibility()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Columns(10, 20).Hide();
        ws.CopyTo("Sheet2");
        Assert.IsTrue(wb.Worksheet("Sheet2").Column(10).IsHidden);
    }

    [Test]
    public void CopyRowVisibility()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Rows(2, 5).Hide();
        ws.CopyTo("Sheet2");
        Assert.IsTrue(wb.Worksheet("Sheet2").Row(4).IsHidden);
    }

    [Test]
    public void DeletingSheets1()
    {
        XLWorkbook wb = new();
        wb.Worksheets.Add("Sheet3");
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet1", 1);

        wb.Worksheet("Sheet3").Delete();

        Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
        Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
        Assert.AreEqual(2, wb.Worksheets.Count);
    }

    [Test]
    public void InsertingSheets1()
    {
        XLWorkbook wb = new();
        wb.Worksheets.Add("Sheet1");
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet3");

        Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
        Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
        Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
    }

    [Test]
    public void InsertingSheets2()
    {
        XLWorkbook wb = new();
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet1", 1);
        wb.Worksheets.Add("Sheet3");

        Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
        Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
        Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
    }

    [Test]
    public void InsertingSheets3()
    {
        XLWorkbook wb = new();
        wb.Worksheets.Add("Sheet3");
        wb.Worksheets.Add("Sheet2", 1);
        wb.Worksheets.Add("Sheet1", 1);

        Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
        Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
        Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
    }

    [Test]
    public void InsertingSheets4()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add();

        Assert.AreEqual("Sheet1", ws1.Name);
        ws1.Name = "shEEt1";

        IXLWorksheet ws2 = wb.Worksheets.Add();
        Assert.AreEqual("Sheet2", ws2.Name);

        wb.Worksheets.Add("SHEET4");

        Assert.AreEqual("Sheet5", wb.Worksheets.Add().Name);
        Assert.AreEqual("Sheet6", wb.Worksheets.Add().Name);

        wb.Worksheets.Add(1);

        Assert.AreEqual("Sheet7", wb.Worksheet(1).Name);
    }

    [Test]
    public void SheetIdIsNotReused()
    {
        using XLWorkbook wb = new();
        XLWorksheet ws1 = (XLWorksheet)wb.AddWorksheet();
        XLWorksheet ws2 = (XLWorksheet)wb.AddWorksheet();
        XLWorksheet ws3 = (XLWorksheet)wb.AddWorksheet();

        Assert.AreEqual(1, ws1.SheetId);
        Assert.AreEqual(2, ws2.SheetId);
        Assert.AreEqual(3, ws3.SheetId);

        ws3.Delete();
        XLWorksheet ws4 = (XLWorksheet)wb.AddWorksheet();
        Assert.AreEqual(4, ws4.SheetId);
    }

    [Test]
    public void AddingDuplicateSheetNameThrowsException()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws;
            ws = wb.AddWorksheet("Sheet1");

            Assert.Throws<ArgumentException>(() => wb.AddWorksheet("Sheet1"));

            //Sheet names are case insensitive
            Assert.Throws<ArgumentException>(() => wb.AddWorksheet("sheet1"));
        }
    }

    [Test]
    public void MergedRanges()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("A1:B2").Merge();
        ws.Range("C1:D3").Merge();
        ws.Range("D2:E2").Merge();

        Assert.AreEqual(2, ws.MergedRanges.Count);
        Assert.AreEqual("A1:B2", ws.MergedRanges.First().RangeAddress.ToStringRelative());
        Assert.AreEqual("D2:E2", ws.MergedRanges.Last().RangeAddress.ToStringRelative());

        Assert.AreEqual("A1:B2", ws.Cell("A2").MergedRange().RangeAddress.ToStringRelative());
        Assert.AreEqual("D2:E2", ws.Cell("D2").MergedRange().RangeAddress.ToStringRelative());

        Assert.AreEqual(null, ws.Cell("Z10").MergedRange());
    }

    [Test]
    public void RowCountTime()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        DateTime start = DateTime.Now;
        ws.RowCount();
        DateTime end = DateTime.Now;
        Assert.IsTrue((end - start).TotalMilliseconds < 500);
    }

    [Test]
    public void SheetsWithCommas()
    {
        using (XLWorkbook wb = new())
        {
            string sourceSheetName = "Sheet1, Sheet3";
            IXLWorksheet ws = wb.Worksheets.Add(sourceSheetName);
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("B2").Value = 3;

            ws = wb.Worksheets.Add("Formula");
            ws.FirstCell().FormulaA1 = string.Format(
                "=SUM('{0}'!A1:A2,'{0}'!B1:B2)",
                sourceSheetName
            );

            XLCellValue value = ws.FirstCell().Value;
            Assert.AreEqual(6, value);
        }
    }

    [Test]
    public void CanRenameWorksheet()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
            IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");

            ws1.Name = "New sheet name";
            Assert.AreEqual("New sheet name", ws1.Name);

            ws2.Name = "sheet2";
            Assert.AreEqual("sheet2", ws2.Name);

            Assert.Throws<ArgumentException>(() => ws1.Name = "SHEET2");
        }
    }

    [Test]
    public void TryGetWorksheet()
    {
        using (XLWorkbook wb = new())
        {
            wb.AddWorksheet("Sheet1");
            wb.AddWorksheet("Sheet2");

            IXLWorksheet ws;
            Assert.IsTrue(wb.Worksheets.TryGetWorksheet("Sheet1", out ws));
            Assert.IsTrue(wb.Worksheets.TryGetWorksheet("sheet1", out ws));
            Assert.IsTrue(wb.Worksheets.TryGetWorksheet("sHEeT1", out ws));
            Assert.IsFalse(wb.Worksheets.TryGetWorksheet("Sheeeet2", out ws));

            Assert.IsTrue(wb.TryGetWorksheet("Sheet1", out ws));
            Assert.IsTrue(wb.TryGetWorksheet("sheet1", out ws));
            Assert.IsTrue(wb.TryGetWorksheet("sHEeT1", out ws));
            Assert.IsFalse(wb.TryGetWorksheet("Sheeeet2", out ws));
        }
    }

    [Test]
    public void HideWorksheet()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                wb.Worksheets.Add("VisibleSheet");
                wb.Worksheets.Add("HiddenSheet").Hide();
                wb.SaveAs(ms);
            }

            // unhide the hidden sheet
            using (XLWorkbook wb = new(ms))
            {
                Assert.AreEqual(
                    XLWorksheetVisibility.Visible,
                    wb.Worksheet("VisibleSheet").Visibility
                );
                Assert.AreEqual(
                    XLWorksheetVisibility.Hidden,
                    wb.Worksheet("HiddenSheet").Visibility
                );

                IXLWorksheet ws = wb.Worksheet("HiddenSheet");
                ws.Unhide().Name = "NoAlsoVisible";

                Assert.AreEqual(XLWorksheetVisibility.Visible, ws.Visibility);

                wb.Save();
            }

            using (XLWorkbook wb = new(ms))
            {
                Assert.AreEqual(
                    XLWorksheetVisibility.Visible,
                    wb.Worksheet("VisibleSheet").Visibility
                );
                Assert.AreEqual(
                    XLWorksheetVisibility.Visible,
                    wb.Worksheet("NoAlsoVisible").Visibility
                );
            }
        }
    }

    [Test]
    public void CanCopySheetsWithAllAnchorTypes()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            ws.CopyTo("Copy1");

            IXLWorksheet ws2 = wb.Worksheets.Skip(1).First();
            ws2.CopyTo("Copy2");

            IXLWorksheet ws3 = wb.Worksheets.Skip(2).First();
            ws3.CopyTo("Copy3");

            IXLWorksheet ws4 = wb.Worksheets.Skip(3).First();
            ws3.CopyTo("Copy4");
        }
    }

    [Test]
    public void CannotCopyDeletedWorksheet()
    {
        using (XLWorkbook wb = new())
        {
            wb.AddWorksheet("Sheet1");
            IXLWorksheet ws = wb.AddWorksheet("Sheet2");

            ws.Delete();
            Assert.Throws<InvalidOperationException>(() => ws.CopyTo("Copy of Sheet2"));
        }
    }

    [Test]
    public void WorksheetNameCannotStartWithApostrophe()
    {
        string title = "'StartsWithApostrophe";
        TestDelegate addWorksheet = () =>
        {
            using (XLWorkbook wb = new())
            {
                wb.Worksheets.Add(title);
            }
        };

        Assert.Throws(typeof(ArgumentException), addWorksheet);
    }

    [Test]
    public void WorksheetNameCannotEndWithApostrophe()
    {
        string title = "EndsWithApostrophe'";
        TestDelegate addWorksheet = () =>
        {
            using (XLWorkbook wb = new())
            {
                wb.Worksheets.Add(title);
            }
        };

        Assert.Throws(typeof(ArgumentException), addWorksheet);
    }

    [Test]
    public void WorksheetNameCannotBeEmpty()
    {
        Assert.Throws<ArgumentException>(() => new XLWorkbook().AddWorksheet(" "));
    }

    [TestCaseSource(nameof(illegalWorksheetCharacters))]
    public void WorksheetNameCannotContainIllegalCharacters(char c)
    {
        string proposedName = $"Sheet{c}Name";
        Assert.Throws<ArgumentException>(() => new XLWorkbook().AddWorksheet(proposedName));
    }

    [Test]
    public void WorksheetNameCanContainApostrophe()
    {
        string title = "With'Apostrophe";
        string savedTitle = "";
        TestDelegate saveAndOpenWorkbook = () =>
        {
            using (MemoryStream ms = new())
            {
                using (XLWorkbook wb = new())
                {
                    wb.Worksheets.Add(title);
                    wb.Worksheets.First().Cell(1, 1).FormulaA1 = $"{title}!A2";
                    wb.SaveAs(ms);
                }

                using (XLWorkbook wb = new(ms))
                {
                    savedTitle = wb.Worksheets.First().Name;
                }
            }
        };

        Assert.DoesNotThrow(saveAndOpenWorkbook);
        Assert.AreEqual(title, savedTitle);
    }

    [Test]
    public void CopyWorksheetPreservesContents()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A1").Value = "A1 value";
            ws1.Cell("A2").Value = 100;
            ws1.Cell("D4").Value = new DateTime(2018, 5, 1);

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual("A1 value", ws2.Cell("A1").Value);
            Assert.AreEqual(100, ws2.Cell("A2").Value);
            Assert.AreEqual(new DateTime(2018, 5, 1), ws2.Cell("D4").Value);
        }
    }

    [Test]
    public void CopyWorksheetPreservesFormulae()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A1").FormulaA1 = "10*10";
            ws1.Cell("A2").FormulaA1 = "A1 * 2";

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual("10*10", ws2.Cell("A1").FormulaA1);
            Assert.AreEqual("A1 * 2", ws2.Cell("A2").FormulaA1);
        }
    }

    [Test]
    public void CopyWorksheetPreservesRowHeights()
    {
        using (XLWorkbook wb1 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");
            using (XLWorkbook wb2 = new())
            {
                ws1.RowHeight = 55;
                ws1.Row(2).Height = 0;
                ws1.Row(3).Height = 20;

                IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.RowHeight, ws2.RowHeight);
                for (int i = 1; i <= 3; i++)
                {
                    Assert.AreEqual(ws1.Row(i).Height, ws2.Row(i).Height);
                }
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesColumnWidths()
    {
        using (XLWorkbook wb1 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");
            using (XLWorkbook wb2 = new())
            {
                ws1.ColumnWidth = 160;
                ws1.Column(2).Width = 0;
                ws1.Column(3).Width = 240;

                IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.ColumnWidth, ws2.ColumnWidth);
                for (int i = 1; i <= 3; i++)
                {
                    Assert.AreEqual(ws1.Column(i).Width, ws2.Column(i).Width);
                }
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesMergedCells()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A:A").Merge();
            ws1.Range("B1:C2").Merge();

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(ws1.MergedRanges.Count, ws2.MergedRanges.Count);
            for (int i = 0; i < ws1.MergedRanges.Count; i++)
            {
                Assert.AreEqual(
                    ws1.MergedRanges.ElementAt(i).RangeAddress.ToString(),
                    ws2.MergedRanges.ElementAt(i).RangeAddress.ToString()
                );
            }
        }
    }

    [Test]
    public void CopySheetAcrossWorkbooksPreservesDefinedNames()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
            ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(ws1.DefinedNames.Count(), ws2.DefinedNames.Count());
            for (int i = 0; i < ws1.DefinedNames.Count(); i++)
            {
                IXLDefinedName nr1 = ws1.DefinedNames.ElementAt(i);
                IXLDefinedName nr2 = ws2.DefinedNames.ElementAt(i);
                Assert.AreEqual(nr1.Ranges.ToString(), nr2.Ranges.ToString());
                Assert.AreEqual(nr1.Scope, nr2.Scope);
                Assert.AreEqual(nr1.Name, nr2.Name);
                Assert.AreEqual(nr1.Visible, nr2.Visible);
                Assert.AreEqual(nr1.Comment, nr2.Comment);
            }
        }
    }

    [Test]
    public void CopyingSheetInsideWorkbookMakesCopiesOfSheetScopedDefinedNames()
    {
        using (XLWorkbook wb1 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
            ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

            IXLWorksheet ws2 = ws1.CopyTo("Copy");

            Assert.AreEqual(ws1.DefinedNames.Count(), ws2.DefinedNames.Count());
            for (int i = 0; i < ws1.DefinedNames.Count(); i++)
            {
                IXLDefinedName nr1 = ws1.DefinedNames.ElementAt(i);
                IXLDefinedName nr2 = ws2.DefinedNames.ElementAt(i);

                Assert.AreEqual(XLScope.Worksheet, nr2.Scope);

                Assert.AreEqual(nr1.Ranges.ToString(), nr2.Ranges.ToString());
                Assert.AreEqual(nr1.Name, nr2.Name);
                Assert.AreEqual(nr1.Visible, nr2.Visible);
                Assert.AreEqual(nr1.Comment, nr2.Comment);
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesStyles()
    {
        using (MemoryStream ms = new())
        using (XLWorkbook wb1 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws1.Range("A1:B2").Style.Font.FontSize = 25;
            ws1.Cell("C3").Style.Fill.BackgroundColor = XLColor.Red;
            ws1.Cell("C4").Style.Fill.BackgroundColor = XLColor.AliceBlue;
            ws1.Cell("C4").Value = "Non empty";

            using (XLWorkbook wb2 = new())
            {
                IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");
                AssertStylesAreEqual(ws1, ws2);
                wb2.SaveAs(ms);
            }

            using (XLWorkbook wb2 = new(ms))
            {
                IXLWorksheet ws2 = wb2.Worksheet("Copy");
                AssertStylesAreEqual(ws1, ws2);
            }
        }

        void AssertStylesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            Assert.That(
                ((XLWorksheet)ws1).FormatValue,
                Is.Not.Null.And.EqualTo(((XLWorksheet)ws2).FormatValue),
                "Worksheet styles differ"
            );
            IXLCells cellsUsed = ws1.Range(ws1.FirstCell(), ws1.LastCellUsed()).Cells();
            foreach (IXLCell cell in cellsUsed)
            {
                XLCellFormatValue? style1 = ((XLCell)cell).FormatValue;
                XLCellFormatValue? style2 = ((XLCell)ws2.Cell(cell.Address.ToString())).FormatValue;
                Assert.AreEqual(style1, style2, $"Cell {cell.Address} styles differ");
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesConditionalFormats()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A:A")
                .AddConditionalFormat()
                .WhenContains("0")
                .Fill.SetBackgroundColor(XLColor.Red);
            IXLConditionalFormat cf = ws1.Range("B1:C2").AddConditionalFormat();
            cf.Ranges = ws1.Ranges("B1:C2,D4:D5");
            cf.WhenEqualOrGreaterThan(100).Font.SetBold();

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(ws1.ConditionalFormats.Count(), ws2.ConditionalFormats.Count());
            for (int i = 0; i < ws1.ConditionalFormats.Count(); i++)
            {
                IXLConditionalFormat original = ws1.ConditionalFormats.ElementAt(i);
                IXLConditionalFormat copy = ws2.ConditionalFormats.ElementAt(i);
                Assert.AreEqual(original.Ranges.ToSpaceList(), copy.Ranges.ToSpaceList());
                Assert.AreEqual(
                    ((XLConditionalFormat)original).FormatValue,
                    ((XLConditionalFormat)copy).FormatValue
                );
                Assert.AreEqual(
                    original.Values.Single().Value.Value,
                    copy.Values.Single().Value.Value
                );
            }

            // Make sure the copy can be saved
            wb2.SaveAs(new MemoryStream());
        }
    }

    [Test]
    public void CopyWorksheetPreservesTables()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A2").Value = "Name";
            ws1.Cell("B2").Value = "Count";
            ws1.Cell("A3").Value = "John Smith";
            ws1.Cell("B3").Value = 50;
            ws1.Cell("A4").Value = "Ivan Ivanov";
            ws1.Cell("B4").Value = 40;
            IXLTable table1 = ws1.Range("A2:B4").CreateTable("Test table 1");
            table1
                .SetShowAutoFilter(true)
                .SetShowTotalsRow(true)
                .SetEmphasizeFirstColumn(true)
                .SetShowColumnStripes(true)
                .SetShowRowStripes(true);
            table1.Theme = XLTableTheme.TableStyleDark8;
            table1.Field(1).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(ws1.Tables.Count(), ws2.Tables.Count());
            for (int i = 0; i < ws1.Tables.Count(); i++)
            {
                IXLTable original = ws1.Tables.ElementAt(i);
                IXLTable copy = ws2.Tables.ElementAt(i);
                Assert.AreEqual(
                    original.RangeAddress.ToString(XLReferenceStyle.A1, false),
                    copy.RangeAddress.ToString(XLReferenceStyle.A1, false)
                );
                Assert.AreEqual(original.Fields.Count(), copy.Fields.Count());
                for (int j = 0; j < original.Fields.Count(); j++)
                {
                    IXLTableField originalField = original.Fields.ElementAt(j);
                    IXLTableField copyField = copy.Fields.ElementAt(j);
                    Assert.AreEqual(originalField.Name, copyField.Name);
                    Assert.AreEqual(originalField.TotalsRowFormulaA1, copyField.TotalsRowFormulaA1);
                    Assert.AreEqual(originalField.TotalsRowFunction, copyField.TotalsRowFunction);
                }

                Assert.AreEqual(original.Name, copy.Name);
                Assert.AreEqual(original.ShowAutoFilter, copy.ShowAutoFilter);
                Assert.AreEqual(original.ShowColumnStripes, copy.ShowColumnStripes);
                Assert.AreEqual(original.ShowHeaderRow, copy.ShowHeaderRow);
                Assert.AreEqual(original.ShowRowStripes, copy.ShowRowStripes);
                Assert.AreEqual(original.ShowTotalsRow, copy.ShowTotalsRow);
                Assert.AreEqual(original.Theme, copy.Theme);
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesDataValidation()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            IXLDataValidation dv1 = ws1.Range("A:A").CreateDataValidation();
            dv1.WholeNumber.EqualTo(2);
            dv1.ErrorStyle = XLErrorStyle.Warning;
            dv1.ErrorTitle = "Number out of range";
            dv1.ErrorMessage = "This cell only allows the number 2.";

            IXLDataValidation dv2 = ws1.Ranges("B2:C3,D4:E5").CreateDataValidation();
            dv2.Decimal.GreaterThan(5);
            dv2.ErrorStyle = XLErrorStyle.Stop;
            dv2.ErrorTitle = "Decimal number out of range";
            dv2.ErrorMessage = "This cell only allows decimals greater than 5.";

            IXLDataValidation dv3 = ws1.Cell("D1").CreateDataValidation();
            dv3.TextLength.EqualOrLessThan(10);
            dv3.ErrorStyle = XLErrorStyle.Information;
            dv3.ErrorTitle = "Text length out of range";
            dv3.ErrorMessage = "You entered more than 10 characters.";

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(ws1.DataValidations.Count(), ws2.DataValidations.Count());
            for (int i = 0; i < ws1.DataValidations.Count(); i++)
            {
                IXLDataValidation original = ws1.DataValidations.ElementAt(i);
                IXLDataValidation copy = ws2.DataValidations.ElementAt(i);

                string originalRanges = string.Join(
                    ",",
                    original.Ranges.Select(r => r.RangeAddress.ToString())
                );
                string copyRanges = string.Join(
                    ",",
                    original.Ranges.Select(r => r.RangeAddress.ToString())
                );

                Assert.AreEqual(originalRanges, copyRanges);
                Assert.AreEqual(original.AllowedValues, copy.AllowedValues);
                Assert.AreEqual(original.Operator, copy.Operator);
                Assert.AreEqual(original.ErrorStyle, copy.ErrorStyle);
                Assert.AreEqual(original.ErrorTitle, copy.ErrorTitle);
                Assert.AreEqual(original.ErrorMessage, copy.ErrorMessage);
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesPictures()
    {
        using (MemoryStream ms = new())
        using (
            Stream? imageStream = Assembly
                .GetAssembly(typeof(BasicTable))
                .GetManifestResourceStream("XlsxSharp.Examples.Resources.SampleImage.jpg")
        )
        using (XLWorkbook wb1 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            IXLPicture picture = ws1.AddPicture(imageStream, "MyPicture")
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50)
                .WithSize(200, 200);

            using (XLWorkbook wb2 = new())
            {
                IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");
                AssertPicturesAreEqual(ws1, ws2);
                wb2.SaveAs(ms);
            }

            using (XLWorkbook wb2 = new(ms))
            {
                IXLWorksheet ws2 = wb2.Worksheet("Copy");
                AssertPicturesAreEqual(ws1, ws2);
            }
        }

        void AssertPicturesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            Assert.AreEqual(ws1.Pictures.Count(), ws2.Pictures.Count());

            for (int i = 0; i < ws1.Pictures.Count(); i++)
            {
                IXLPicture original = ws1.Pictures.ElementAt(i);
                IXLPicture copy = ws2.Pictures.ElementAt(i);
                Assert.AreEqual(ws2, copy.Worksheet);

                Assert.AreEqual(original.Format, copy.Format);
                Assert.AreEqual(original.Height, copy.Height);
                Assert.AreEqual(original.Id, copy.Id);
                Assert.AreEqual(original.Left, copy.Left);
                Assert.AreEqual(original.Name, copy.Name);
                Assert.AreEqual(original.Placement, copy.Placement);
                Assert.AreEqual(original.Top, copy.Top);
                Assert.AreEqual(
                    original.TopLeftCell.Address.ToString(),
                    copy.TopLeftCell.Address.ToString()
                );
                Assert.AreEqual(original.Width, copy.Width);
                Assert.AreEqual(
                    original.ImageStream.ToArray(),
                    copy.ImageStream.ToArray(),
                    "Image streams differ"
                );
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesPivotTables()
    {
        using (MemoryStream ms = new())
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws1 = wb.Worksheet("pvt1");
            IXLWorksheet copyOfws1 = ws1.CopyTo("CopyOfPvt1");

            AssertPivotTablesAreEqual(ws1, copyOfws1);

            using (XLWorkbook wb2 = new())
            {
                // We need to  copy the source too. Cross workbook references don't work yet.
                wb.Worksheet("PastrySalesData").CopyTo(wb2);
                IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");
                AssertPivotTablesAreEqual(ws1, ws2);
                wb2.SaveAs(ms);
            }

            using (XLWorkbook wb2 = new(ms))
            {
                IXLWorksheet ws2 = wb2.Worksheet("Copy");
                AssertPivotTablesAreEqual(ws1, ws2);
            }
        }

        void AssertPivotTablesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            Assert.AreEqual(ws1.PivotTables.Count(), ws2.PivotTables.Count());

            PivotTableComparer comparer = new();

            for (int i = 0; i < ws1.PivotTables.Count(); i++)
            {
                XLPivotTable original = ws1.PivotTables.ElementAt(i).CastTo<XLPivotTable>();
                XLPivotTable copy = ws2.PivotTables.ElementAt(i).CastTo<XLPivotTable>();

                Assert.AreEqual(ws2, copy.Worksheet);

                Assert.IsTrue(comparer.Equals(original, copy));
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesSelectedRanges()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.SelectedRanges.RemoveAll();
            ws1.SelectedRanges.Add(ws1.Range("E12:H20"));
            ws1.SelectedRanges.Add(ws1.Range("B:B"));
            ws1.SelectedRanges.Add(ws1.Range("3:6"));

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(ws1.SelectedRanges.Count, ws2.SelectedRanges.Count);
            for (int i = 0; i < ws1.SelectedRanges.Count; i++)
            {
                Assert.AreEqual(
                    ws1.SelectedRanges.ElementAt(i).RangeAddress.ToString(),
                    ws2.SelectedRanges.ElementAt(i).RangeAddress.ToString()
                );
            }
        }
    }

    [Test]
    public void CopyWorksheetPreservesPageSetup()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.PageSetup.AddHorizontalPageBreak(15);
            ws1.PageSetup.AddVerticalPageBreak(5);
            ws1.PageSetup.SetBlackAndWhite()
                .SetCenterHorizontally()
                .SetCenterVertically()
                .SetFirstPageNumber(200)
                .SetPageOrientation(XLPageOrientation.Landscape)
                .SetPaperSize(XLPaperSize.A5Paper)
                .SetScale(89)
                .SetShowGridlines()
                .SetHorizontalDpi(200)
                .SetVerticalDpi(300)
                .SetPagesTall(5)
                .SetPagesWide(2)
                .SetColumnsToRepeatAtLeft(1, 3);
            ws1.PageSetup.PrintAreas.Clear();
            ws1.PageSetup.PrintAreas.Add("A1:Z200");
            ws1.PageSetup.Margins.SetBottom(5)
                .SetTop(6)
                .SetLeft(7)
                .SetRight(8)
                .SetFooter(9)
                .SetHeader(10);
            ws1.PageSetup.Header.Left.AddText(XLHFPredefinedText.FullPath, XLHFOccurrence.AllPages);
            ws1.PageSetup.Footer.Right.AddText(
                XLHFPredefinedText.PageNumber,
                XLHFOccurrence.OddPages
            );

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(
                ws1.PageSetup.FirstRowToRepeatAtTop,
                ws2.PageSetup.FirstRowToRepeatAtTop
            );
            Assert.AreEqual(ws1.PageSetup.LastRowToRepeatAtTop, ws2.PageSetup.LastRowToRepeatAtTop);
            Assert.AreEqual(
                ws1.PageSetup.FirstColumnToRepeatAtLeft,
                ws2.PageSetup.FirstColumnToRepeatAtLeft
            );
            Assert.AreEqual(
                ws1.PageSetup.LastColumnToRepeatAtLeft,
                ws2.PageSetup.LastColumnToRepeatAtLeft
            );
            Assert.AreEqual(ws1.PageSetup.PageOrientation, ws2.PageSetup.PageOrientation);
            Assert.AreEqual(ws1.PageSetup.PagesWide, ws2.PageSetup.PagesWide);
            Assert.AreEqual(ws1.PageSetup.PagesTall, ws2.PageSetup.PagesTall);
            Assert.AreEqual(ws1.PageSetup.Scale, ws2.PageSetup.Scale);
            Assert.AreEqual(ws1.PageSetup.HorizontalDpi, ws2.PageSetup.HorizontalDpi);
            Assert.AreEqual(ws1.PageSetup.VerticalDpi, ws2.PageSetup.VerticalDpi);
            Assert.AreEqual(ws1.PageSetup.FirstPageNumber, ws2.PageSetup.FirstPageNumber);
            Assert.AreEqual(ws1.PageSetup.CenterHorizontally, ws2.PageSetup.CenterHorizontally);
            Assert.AreEqual(ws1.PageSetup.CenterVertically, ws2.PageSetup.CenterVertically);
            Assert.AreEqual(ws1.PageSetup.PaperSize, ws2.PageSetup.PaperSize);
            Assert.AreEqual(ws1.PageSetup.Margins.Bottom, ws2.PageSetup.Margins.Bottom);
            Assert.AreEqual(ws1.PageSetup.Margins.Top, ws2.PageSetup.Margins.Top);
            Assert.AreEqual(ws1.PageSetup.Margins.Left, ws2.PageSetup.Margins.Left);
            Assert.AreEqual(ws1.PageSetup.Margins.Right, ws2.PageSetup.Margins.Right);
            Assert.AreEqual(ws1.PageSetup.Margins.Footer, ws2.PageSetup.Margins.Footer);
            Assert.AreEqual(ws1.PageSetup.Margins.Header, ws2.PageSetup.Margins.Header);
            Assert.AreEqual(ws1.PageSetup.ScaleHFWithDocument, ws2.PageSetup.ScaleHFWithDocument);
            Assert.AreEqual(ws1.PageSetup.AlignHFWithMargins, ws2.PageSetup.AlignHFWithMargins);
            Assert.AreEqual(ws1.PageSetup.ShowGridlines, ws2.PageSetup.ShowGridlines);
            Assert.AreEqual(
                ws1.PageSetup.ShowRowAndColumnHeadings,
                ws2.PageSetup.ShowRowAndColumnHeadings
            );
            Assert.AreEqual(ws1.PageSetup.BlackAndWhite, ws2.PageSetup.BlackAndWhite);
            Assert.AreEqual(ws1.PageSetup.DraftQuality, ws2.PageSetup.DraftQuality);
            Assert.AreEqual(ws1.PageSetup.PageOrder, ws2.PageSetup.PageOrder);
            Assert.AreEqual(ws1.PageSetup.ShowComments, ws2.PageSetup.ShowComments);
            Assert.AreEqual(ws1.PageSetup.PrintErrorValue, ws2.PageSetup.PrintErrorValue);

            Assert.AreEqual(ws1.PageSetup.PrintAreas.Count(), ws2.PageSetup.PrintAreas.Count());

            Assert.AreEqual(
                ws1.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages),
                ws2.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages)
            );
            Assert.AreEqual(
                ws1.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages),
                ws2.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages)
            );
        }
    }

    [Test]
    public void CopyWorksheetPreservesSparklineGroups()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");
            IXLSparklineGroup original = ws1
                .SparklineGroups.Add("A1:A10", "D1:Z10")
                .SetDateRange(ws1.Range("D11:Z11"))
                .SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues.Zero)
                .SetDisplayHidden(true)
                .SetLineWeight(1.5)
                .SetShowMarkers(XLSparklineMarkers.All)
                .SetStyle(XLSparklineTheme.Colorful3)
                .SetType(XLSparklineType.Column);

            original.HorizontalAxis.SetColor(XLColor.Blue).SetRightToLeft(true).SetVisible(true);

            original.VerticalAxis.SetManualMin(-100.0).SetManualMax(100.0);

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual(1, ws2.SparklineGroups.Count());
            IXLSparklineGroup copy = ws2.SparklineGroups.Single();

            Assert.AreEqual(original.Count(), copy.Count());
            for (int i = 0; i < original.Count(); i++)
            {
                Assert.AreSame(ws2, copy.ElementAt(i).Location.Worksheet);
                Assert.AreSame(ws2, copy.ElementAt(i).SourceData.Worksheet);
                Assert.AreEqual(
                    original.ElementAt(i).Location.Address.ToString(),
                    copy.ElementAt(i).Location.Address.ToString()
                );
                Assert.AreEqual(
                    original.ElementAt(i).SourceData.RangeAddress.ToString(),
                    copy.ElementAt(i).SourceData.RangeAddress.ToString()
                );
            }

            Assert.AreEqual(
                original.DateRange.RangeAddress.ToString(),
                copy.DateRange.RangeAddress.ToString()
            );
            Assert.AreSame(ws2, copy.DateRange.Worksheet);

            Assert.AreEqual(original.DisplayEmptyCellsAs, copy.DisplayEmptyCellsAs);
            Assert.AreEqual(original.DisplayHidden, copy.DisplayHidden);
            Assert.AreEqual(original.LineWeight, copy.LineWeight, XLHelper.Epsilon);
            Assert.AreEqual(original.ShowMarkers, copy.ShowMarkers);
            Assert.AreEqual(original.Style, copy.Style);
            Assert.AreNotSame(original.Style, copy.Style);
            Assert.AreEqual(original.Type, copy.Type);

            Assert.AreEqual(original.HorizontalAxis.Color, copy.HorizontalAxis.Color);
            Assert.AreEqual(original.HorizontalAxis.DateAxis, copy.HorizontalAxis.DateAxis);
            Assert.AreEqual(original.HorizontalAxis.IsVisible, copy.HorizontalAxis.IsVisible);
            Assert.AreEqual(original.HorizontalAxis.RightToLeft, copy.HorizontalAxis.RightToLeft);

            Assert.AreEqual(original.VerticalAxis.ManualMax, copy.VerticalAxis.ManualMax);
            Assert.AreEqual(original.VerticalAxis.ManualMin, copy.VerticalAxis.ManualMin);
            Assert.AreEqual(original.VerticalAxis.MaxAxisType, copy.VerticalAxis.MaxAxisType);
            Assert.AreEqual(original.VerticalAxis.MinAxisType, copy.VerticalAxis.MinAxisType);
        }
    }

    [Test, Ignore("Muted until #836 is fixed")]
    public void CopyWorksheetChangesAbsoluteReferencesInFormulae()
    {
        using (XLWorkbook wb1 = new())
        using (XLWorkbook wb2 = new())
        {
            IXLWorksheet ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A1").FormulaA1 = "10*10";
            ws1.Cell("A2").FormulaA1 = "Original!A1 * 3";

            IXLWorksheet ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.AreEqual("Copy!A1 * 3", ws2.Cell("A2").FormulaA1);
        }
    }

    [Test]
    public void RenameSheetsChangesSheetReferencesInFormulas()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Original");

        ws.Cell("A1").FormulaA1 = "10*10";
        ws.Cell("A2").FormulaA1 = "Original!A1 * 3";
        _ = ws.Cell("A2").Value;

        ws.Name = "Renamed";

        Assert.AreEqual("Renamed!A1 * 3", ws.Cell("A2").FormulaA1);
        Assert.True(ws.Cell("A2").NeedsRecalculation);
        Assert.AreEqual(300, ws.Cell("A2").Value);
    }

    [Test]
    public void RangesFromDeletedWorksheetContainREF()
    {
        using (XLWorkbook wb1 = new())
        {
            wb1.Worksheets.Add("Sheet1");
            IXLWorksheet ws2 = wb1.Worksheets.Add("Sheet2");
            IXLRange range = ws2.Range("A1:B2");

            ws2.Delete();

            Assert.AreEqual("#REF!A1:B2", range.RangeAddress.ToString());
        }
    }

    [Test]
    public void InvalidRowAndColumnIndices()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(XLHelper.MaxRowNumber + 1));

            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Column(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                ws.Column(XLHelper.MaxColumnNumber + 1)
            );
        }
    }

    [Test]
    public void InvalidSelectedRangeExcluded()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLRange range1 = ws.Range("B2:C2");
            IXLRange range2 = ws.Range("B4:C4");
            ws.SelectedRanges.Clear();

            ws.SelectedRanges.Add(range1);
            ws.SelectedRanges.Add(range2);

            ws.Row(4).Delete();

            Assert.IsFalse(range2.RangeAddress.IsValid);
            Assert.AreEqual(range1, ws.SelectedRanges.Single());
        }
    }

    [Test]
    public void InsertColumnsDoesNotIncreaseCellsCount()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.Cell("A1").SetValue(1);
            ws.Cell("AAA50").SetValue(1);
            int originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            ws.Column(1).InsertColumnsBefore(1);

            Assert.AreEqual(
                originalCount,
                ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()
            );
        }
    }

    [Test]
    public void InsertRowsDoesNotIncreaseCellsCount()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.Cell("A1").SetValue(1);
            ws.Cell("AAA500").SetValue(1);
            int originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            ws.Row(1).InsertRowsAbove(1);

            Assert.AreEqual(
                originalCount,
                ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()
            );
        }
    }

    [Test]
    public void InsertCellsBeforeDoesNotIncreaseCellsCount()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLCell a1 = ws.Cell("A1").SetValue(1);
            ws.Cell("AAA50").SetValue(1);
            int originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            a1.InsertCellsBefore(1);

            Assert.AreEqual(
                originalCount,
                ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()
            );
        }
    }

    [Test]
    public void InsertCellsAboveDoesNotIncreaseCellsCount()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLCell a1 = ws.Cell("A1").SetValue(1);
            ws.Cell("AAA500").SetValue(1);
            int originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            a1.InsertCellsAbove(1);

            Assert.AreEqual(
                originalCount,
                ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()
            );
        }
    }

    [Test]
    public void CellsShiftedTooFarRightArePurged()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLCell a1 = ws.Cell("A1").SetValue(1);
            ws.Cell(1, XLHelper.MaxColumnNumber).SetValue(1);
            ws.Cell(2, XLHelper.MaxColumnNumber).SetValue(1);

            a1.InsertCellsBefore(1);

            Assert.AreEqual(2, ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count());
            ws.Column(1).InsertColumnsBefore(1);
            Assert.AreEqual(1, ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count());
        }
    }

    [Test]
    public void CellsShiftedTooFarDownArePurged()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLCell a1 = ws.Cell("A1").SetValue(1);
            ws.Cell(XLHelper.MaxRowNumber, 1).SetValue(1);
            ws.Cell(XLHelper.MaxRowNumber, 2).SetValue(1);

            a1.InsertCellsAbove(1);

            Assert.AreEqual(2, ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count());
            ws.Row(1).InsertRowsAbove(1);
            Assert.AreEqual(1, ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count());
        }
    }

    [Test]
    public void MaxColumnUsedUpdatedWhenColumnDeleted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.Cell("C1").SetValue(1);
            ws.Cell(1, XLHelper.MaxColumnNumber).SetValue(1);

            ws.Column(XLHelper.MaxColumnNumber).Delete();

            Assert.AreEqual(3, ((XLWorksheet)ws).Internals.CellsCollection.MaxColumnUsed);
        }
    }

    [Test]
    public void MaxRowUsedUpdatedWhenRowDeleted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.Cell("A3").SetValue(1);
            ws.Cell(XLHelper.MaxRowNumber, 1).SetValue(1);

            ws.Row(XLHelper.MaxRowNumber).Delete();

            Assert.AreEqual(3, ((XLWorksheet)ws).Internals.CellsCollection.MaxRowUsed);
        }
    }

    [Test]
    public void ChangeColumnStyleFirst()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("ColumnFirst");

            ws.Column(2).Style.Font.SetBold(true);
            ws.Row(2).Style.Font.SetItalic(true);

            Assert.IsTrue(ws.Cell("B2").Style.Font.Bold);
            Assert.IsTrue(ws.Cell("B2").Style.Font.Italic);
        }
    }

    [Test]
    public void ChangeRowStyleFirst()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("RowFirst");

            ws.Row(2).Style.Font.SetItalic(true);
            ws.Column(2).Style.Font.SetBold(true);

            Assert.IsTrue(ws.Cell("B2").Style.Font.Bold);
            Assert.IsTrue(ws.Cell("B2").Style.Font.Italic);
        }
    }

    [Test]
    public void SelectedTabIsActiveWhenInsertBefore()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws1 = wb.AddWorksheet();
                ws1.TabSelected = true;
                IXLWorksheet ws2 = wb.Worksheets.Add(1);
                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws1 = wb.Worksheets.First();
                IXLWorksheet ws2 = wb.Worksheets.Last();

                Assert.IsFalse(ws1.TabActive);
                Assert.IsFalse(ws1.TabSelected);
                Assert.IsTrue(ws2.TabActive);
                Assert.IsTrue(ws2.TabSelected);
            }
        }
    }

    [TestCase("noactive_noselected.xlsx")]
    [TestCase("noactive_twoselected.xlsx")]
    [TestCase("noactive_negativeId.xlsx")]
    public void FirstSheetIsActiveWhenNotSpecified(string fileName)
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Other\NoActiveSheet\" + fileName)
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            Assert.IsTrue(wb.Worksheets.First().TabActive);
            Assert.AreEqual(XLWorksheetVisibility.Visible, wb.Worksheets.First().Visibility);
        }
    }

    [TestCase(XLCellsUsedOptions.NormalFormats, 42)]
    [TestCase(XLCellsUsedOptions.Contents, 100)]
    public void FirstColumnUsedReturnsFirstColumnWithUsedCell(
        XLCellsUsedOptions options,
        int expectedColumn
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 42).Style.Fill.SetBackgroundColor(XLColor.Green);
        ws.Cell(1, 100).SetValue(5);

        IXLColumn? column = ws.FirstColumnUsed(options);
        Assert.AreEqual(expectedColumn, column.ColumnNumber());
    }

    [Test]
    public void RecalculateAllFormulasRecalculatesAllFormulasInSheetAndLeavesRestDirty()
    {
        using XLWorkbook wb = new();
        IXLWorksheet sut = wb.AddWorksheet("sut");
        IXLWorksheet other = wb.AddWorksheet("other");

        other.Cell("A1").Value = 7;
        other.Cell("A2").FormulaA1 = "A1+3";
        Assert.AreEqual(10.0, other.Cell("A2").Value);

        // Change the supporting value, but without recalculation of dependent
        // formula, thus the value stays the same.
        other.Cell("A1").Value = 5;

        Assert.True(other.Cell("A2").NeedsRecalculation);
        Assert.AreEqual(10.0, other.Cell("A2").CachedValue);

        // Tested formula depends on a dirty formula from other sheet.
        sut.Cell("A1").FormulaA1 = "other!A2+5";
        sut.Cell("A2").FormulaA1 = "1+2";

        Assert.AreEqual(Blank.Value, sut.Cell("A1").CachedValue);
        Assert.AreEqual(Blank.Value, sut.Cell("A2").CachedValue);

        sut.RecalculateAllFormulas();

        // Formulas in other sheets kept the value - not affected by recalculation of a sut sheet.
        Assert.True(other.Cell("A2").NeedsRecalculation);
        Assert.AreEqual(10.0, other.Cell("A2").CachedValue);

        // Formulas in test sheet were recalculated - they are affected by recalculation of a sut sheet.
        Assert.False(sut.Cell("A1").NeedsRecalculation);
        Assert.AreEqual(15.0, sut.Cell("A1").CachedValue);

        Assert.False(sut.Cell("A2").NeedsRecalculation);
        Assert.AreEqual(3.0, sut.Cell("A2").CachedValue);
    }

    [Test]
    public void CellReturnsCellAtAddressOrWorkbookScopedNamedRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        wb.DefinedNames.Add("test_range", ws.Range(2, 3, 5, 7)); // C2:G5

        IXLCell cellB4 = ws.Cell("B4");
        IXLCell firstCellOfRange = ws.Cell("test_range");

        Assert.AreEqual("B4", cellB4.Address.ToString());
        Assert.AreEqual("C2", firstCellOfRange.Address.ToString());
    }

    [Test]
    public void CellThrowsExceptionWhenAddressIsNotA1AddressOrWorkbookScopedRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        Assert.Throws<ArgumentException>(() => _ = ws.Cell("XFF1"));
        Assert.Throws<ArgumentException>(() => _ = ws.Cell("nonexistent_range"));
    }

    [Test]
    public void RangeReturnsRangeFromA1AddressOrNamedRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        wb.DefinedNames.Add("book_range", ws.Range(2, 3, 5, 7)); // C2:G5
        ws.DefinedNames.Add("sheet_range", ws.Range(1, 2, 3, 4)); // B1:D3

        IXLRange singleCellRange = ws.Range("B4");
        IXLRange areaCellRange = ws.Range("B4:D7");
        IXLRange bookNamedRange = ws.Range("book_range");
        IXLRange sheetNamedRange = ws.Range("sheet_range");

        Assert.AreEqual("B4:B4", singleCellRange.RangeAddress.ToString());
        Assert.AreEqual("B4:D7", areaCellRange.RangeAddress.ToString());
        Assert.AreEqual("$C$2:$G$5", bookNamedRange.RangeAddress.ToString());
        Assert.AreEqual("$B$1:$D$3", sheetNamedRange.RangeAddress.ToString());
    }

    [Test]
    public void RangeThrowsExceptionWhenAddressIsNotA1AddressOrNamedRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        Assert.Throws<ArgumentException>(() => _ = ws.Range("DEAD1"));
        Assert.Throws<ArgumentException>(() => _ = ws.Range("DEAD4:BEEF10"));
        Assert.Throws<ArgumentException>(() => _ = ws.Range("nonexistent_range"));
    }

    [TestCase("Sheet1", "Sheet1")]
    [TestCase("Sheet1", "SHEET1")]
    [TestCase("Baker's Paradise", "BAKER'S PARADISE")]
    [TestCase("XXX''XXX", "XXX''XXX")]
    public void WorksheetByNameReturnsWorksheetWithTheSameCaseInsensitiveName(
        string sheetName,
        string searchedSheetName
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet(sheetName);

        Assert.That(wb.Worksheets.Worksheet(searchedSheetName), Is.SameAs(sheet));
    }

    [Test]
    public void WorksheetByNameThrowsExceptionWhenNoSheetWithNameFound()
    {
        using XLWorkbook wb = new();
        wb.AddWorksheet("Sheet");

        Assert.That(
            () => wb.Worksheets.Worksheet("Nonexistent"),
            Throws.TypeOf<KeyNotFoundException>()
        );
    }

    [TestCase("Sheet1", "Sheet1", true)]
    [TestCase("Sheet1", "SHEET1", true)]
    [TestCase("Sheet1", "Sheet", false)]
    [TestCase("Sheet1", " Sheet1", false)]
    [TestCase("Baker's Paradise", "BAKER'S PARADISE", true)]
    [TestCase("XXX''XXX", "XXX''XXX", true)]
    public void TryGetWorksheetFindsWorksheetWithTheSameCaseInsensitiveName(
        string sheetName,
        string searchedSheetName,
        bool expectedFound
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet(sheetName);

        bool found = wb.Worksheets.TryGetWorksheet(searchedSheetName, out IXLWorksheet? foundSheet);

        Assert.AreEqual(expectedFound, found);
        Assert.That(foundSheet, found ? Is.SameAs(sheet) : Is.Null);
    }
}
