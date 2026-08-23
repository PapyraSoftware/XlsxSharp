using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.Ranges;

public class MergedRangesTests
{
    [Test]
    public void LastCellFromMerge()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");
        ws.Range("B2:D4").Merge();

        string first = ws.FirstCellUsed(XLCellsUsedOptions.All).Address.ToStringRelative();
        string last = ws.LastCellUsed(XLCellsUsedOptions.All).Address.ToStringRelative();

        ClassicAssert.AreEqual("B2", first);
        ClassicAssert.AreEqual("D4", last);
    }

    [Test]
    [Arguments("A1:A2", "A1:A2")]
    [Arguments("A2:B2", "A2:B2")]
    [Arguments("A3:C3", "A3:E3")]
    [Arguments("B4:B6", "B4:B6")]
    [Arguments("C7:D7", "E7:F7")]
    public void MergedRangesShiftedOnColumnInsert(string originalRange, string expectedRange)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            IXLRange range = ws.Range(originalRange).Merge();

            ws.Column(2).InsertColumnsAfter(2);

            IXLRange[] mr = [.. ws.MergedRanges];
            ClassicAssert.AreEqual(1, mr.Length);
            ClassicAssert.AreSame(range, mr.Single());
            ClassicAssert.AreEqual(expectedRange, range.RangeAddress.ToString());
        }
    }

    [Test]
    [Arguments("A1:B1", "A1:B1")]
    [Arguments("B1:B2", "B1:B2")]
    [Arguments("C1:C3", "C1:C5")]
    [Arguments("D2:F2", "D2:F2")]
    [Arguments("G4:G5", "G6:G7")]
    public void MergedRangesShiftedOnRowInsert(string originalRange, string expectedRange)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            IXLRange range = ws.Range(originalRange).Merge();

            ws.Row(2).InsertRowsBelow(2);

            IXLRange[] mr = [.. ws.MergedRanges];
            ClassicAssert.AreEqual(1, mr.Length);
            ClassicAssert.AreSame(range, mr.Single());
            ClassicAssert.AreEqual(expectedRange, range.RangeAddress.ToString());
        }
    }

    [Test]
    [Arguments("A1:A2", true, "A1:A2")]
    [Arguments("A2:B2", true, "A2:A2")]
    [Arguments("A3:C3", true, "A3:B3")]
    [Arguments("B4:B6", false, "")]
    [Arguments("C7:D7", true, "B7:C7")]
    public void MergedRangesShiftedOnColumnDelete(
        string originalRange,
        bool expectedExist,
        string expectedRange
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            IXLRange range = ws.Range(originalRange).Merge();

            ws.Column(2).Delete();

            IXLRange[] mr = [.. ws.MergedRanges];
            if (expectedExist)
            {
                ClassicAssert.AreEqual(1, mr.Length);
                ClassicAssert.AreSame(range, mr.Single());
                ClassicAssert.AreEqual(expectedRange, range.RangeAddress.ToString());
            }
            else
            {
                ClassicAssert.AreEqual(0, mr.Length);
                ClassicAssert.IsFalse(range.RangeAddress.IsValid);
            }
        }
    }

    [Test]
    [Arguments("A1:B1", true, "A1:B1")]
    [Arguments("B1:B2", true, "B1:B1")]
    [Arguments("C1:C3", true, "C1:C2")]
    [Arguments("D2:F2", false, "")]
    [Arguments("G4:G5", true, "G3:G4")]
    public void MergedRangesShiftedOnRowDelete(
        string originalRange,
        bool expectedExist,
        string expectedRange
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            IXLRange range = ws.Range(originalRange).Merge();

            ws.Row(2).Delete();

            IXLRange[] mr = [.. ws.MergedRanges];
            if (expectedExist)
            {
                ClassicAssert.AreEqual(1, mr.Length);
                ClassicAssert.AreSame(range, mr.Single());
                ClassicAssert.AreEqual(expectedRange, range.RangeAddress.ToString());
            }
            else
            {
                ClassicAssert.AreEqual(0, mr.Length);
                ClassicAssert.IsFalse(range.RangeAddress.IsValid);
            }
        }
    }

    [Test]
    public void ShiftRangeRightBreaksMerges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("B4:C5").Merge();
            ws.Range("F2:G3").Merge(); // to be broken
            ws.Range("F4:G5").Merge(); // to be broken
            ws.Range("H1:I2").Merge();
            ws.Range("H5:I6").Merge();

            ws.Range("D3:E4").InsertColumnsAfter(2);

            IXLRange[] mr = [.. ws.MergedRanges];
            ClassicAssert.AreEqual(4, mr.Length);
            ClassicAssert.AreEqual("H1:I2", mr[0].RangeAddress.ToString());
            ClassicAssert.AreEqual("B2:C3", mr[1].RangeAddress.ToString());
            ClassicAssert.AreEqual("B4:C5", mr[2].RangeAddress.ToString());
            ClassicAssert.AreEqual("H5:I6", mr[3].RangeAddress.ToString());
        }
    }

    [Test]
    public void ShiftRangeLeftBreaksMerges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("B4:C5").Merge();
            ws.Range("F2:G3").Merge(); // to be broken
            ws.Range("F4:G5").Merge(); // to be broken
            ws.Range("H1:I2").Merge();
            ws.Range("H5:I6").Merge();

            ws.Range("D3:E4").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            IXLRange[] mr = [.. ws.MergedRanges];
            ClassicAssert.AreEqual(4, mr.Length);
            ClassicAssert.AreEqual("H1:I2", mr[0].RangeAddress.ToString());
            ClassicAssert.AreEqual("B2:C3", mr[1].RangeAddress.ToString());
            ClassicAssert.AreEqual("B4:C5", mr[2].RangeAddress.ToString());
            ClassicAssert.AreEqual("H5:I6", mr[3].RangeAddress.ToString());
        }
    }

    [Test]
    public void RangeShiftDownBreaksMerges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("D2:E3").Merge();
            ws.Range("B6:C7").Merge(); // to be broken
            ws.Range("D6:E7").Merge(); // to be broken
            ws.Range("A8:B9").Merge();
            ws.Range("E8:F9").Merge();

            ws.Range("C4:D5").InsertRowsBelow(2);

            IXLRange[] mr = [.. ws.MergedRanges];
            ClassicAssert.AreEqual(4, mr.Length);
            ClassicAssert.AreEqual("B2:C3", mr[0].RangeAddress.ToString());
            ClassicAssert.AreEqual("D2:E3", mr[1].RangeAddress.ToString());
            ClassicAssert.AreEqual("A8:B9", mr[2].RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:F9", mr[3].RangeAddress.ToString());
        }
    }

    [Test]
    public void RangeShiftUpBreaksMerges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("D2:E3").Merge();
            ws.Range("B6:C7").Merge(); // to be broken
            ws.Range("D6:E7").Merge(); // to be broken
            ws.Range("A8:B9").Merge();
            ws.Range("E8:F9").Merge();

            ws.Range("C4:D5").Delete(XLShiftDeletedCells.ShiftCellsUp);

            IXLRange[] mr = [.. ws.MergedRanges];
            ClassicAssert.AreEqual(4, mr.Length);
            ClassicAssert.AreEqual("B2:C3", mr[0].RangeAddress.ToString());
            ClassicAssert.AreEqual("D2:E3", mr[1].RangeAddress.ToString());
            ClassicAssert.AreEqual("A8:B9", mr[2].RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:F9", mr[3].RangeAddress.ToString());
        }
    }

    [Test]
    public void MergedCellsAcquireFirstCellStyle()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
            ws.Range("A1:A3").Merge();

            ClassicAssert.AreEqual(XLColor.Red, ws.Cell("A1").Style.Fill.BackgroundColor);
            ClassicAssert.AreEqual(XLColor.Red, ws.Cell("A2").Style.Fill.BackgroundColor);
            ClassicAssert.AreEqual(XLColor.Red, ws.Cell("A3").Style.Fill.BackgroundColor);
        }
    }

    [Test]
    public void MergedCellsLooseData()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A3").SetValue(100);
            ws.Range("A1:A3").Merge();

            ClassicAssert.AreEqual(100, ws.Cell("A1").Value);
            ClassicAssert.AreEqual(Blank.Value, ws.Cell("A2").Value);
            ClassicAssert.AreEqual(Blank.Value, ws.Cell("A3").Value);
        }
    }

    [Test]
    public void MergedCellsLooseConditionalFormats()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").AddConditionalFormat().WhenContains("1").Fill.BackgroundColor =
                XLColor.Red;
            ws.Cell("A2").AddConditionalFormat().WhenContains("2").Fill.BackgroundColor =
                XLColor.Yellow;

            ws.Range("A1:A2").Merge();

            ClassicAssert.AreEqual(1, ws.ConditionalFormats.Count());
            ClassicAssert.AreEqual(
                "A1:A1",
                ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString()
            );
        }
    }

    [Test]
    public void MergedCellsLooseDataValidation()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").CreateDataValidation().WholeNumber.Between(1, 2);
            ws.Cell("A2").CreateDataValidation().Date.GreaterThan(new System.DateTime(2018, 1, 1));

            ws.Range("A1:A2").Merge();

            ClassicAssert.IsTrue(ws.Cell("A1").HasDataValidation);
            ClassicAssert.AreEqual("1", ws.Cell("A1").GetDataValidation().MinValue);
            ClassicAssert.AreEqual("2", ws.Cell("A1").GetDataValidation().MaxValue);
            ClassicAssert.IsFalse(ws.Cell("A2").HasDataValidation);
        }
    }

    [Test]
    public void UnmergedCellsPreserveStyle()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLRange range = ws.Range("B2:D4");
            range.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            range
                .Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.DarkBlue)
                .Border.SetInsideBorder(XLBorderStyleValues.Thin)
                .Border.SetInsideBorderColor(XLColor.Pink);
            range.Cells().ForEach(c => c.Value = c.Address.ToString());

            IXLCell firstCell = ws.Cell("B2");
            firstCell
                .Style.Fill.SetBackgroundColor(XLColor.Red)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetBold();

            range.Merge();
            range.Unmerge();

            ClassicAssert.IsTrue(
                range.Cells().All(c => c.Style.Fill.BackgroundColor == XLColor.Red)
            );
            ClassicAssert.IsTrue(
                range.Cells().Where(c => !c.Equals(firstCell)).All(c => c.Value.Equals(Blank.Value))
            );
            ClassicAssert.AreEqual("B2", firstCell.Value);

            ClassicAssert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("B2").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("B2").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("B2").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("B2").Style.Border.LeftBorder
            );

            ClassicAssert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("C2").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("C2").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("C2").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("C2").Style.Border.LeftBorder);

            ClassicAssert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("D2").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("D2").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("D2").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("D2").Style.Border.LeftBorder);

            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("B3").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("B3").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("B3").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("B3").Style.Border.LeftBorder
            );

            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("C3").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("C3").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("C3").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("C3").Style.Border.LeftBorder);

            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("D3").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("D3").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("D3").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("D3").Style.Border.LeftBorder);

            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("B4").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("B4").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("B4").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("B4").Style.Border.LeftBorder
            );

            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("C4").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.None,
                ws.Cell("C4").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("C4").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("C4").Style.Border.LeftBorder);

            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("D4").Style.Border.TopBorder);
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("D4").Style.Border.RightBorder
            );
            ClassicAssert.AreEqual(
                XLBorderStyleValues.Thick,
                ws.Cell("D4").Style.Border.BottomBorder
            );
            ClassicAssert.AreEqual(XLBorderStyleValues.None, ws.Cell("D4").Style.Border.LeftBorder);
        }
    }

    [Test]
    public void MergedRangesCellValuesShouldNotBeSet()
    {
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").Value = 1;
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").Value = 1;
            ws.Cell("B1").FormulaA1 = "SUM(A:A)";
            ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        }

        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge().SetValue(1);
            ws.Cell("B1").FormulaA1 = "SUM(A:A)";
            ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        }
    }

    [Test]
    public void MergedRangesCellFormulasShouldNotBeSet()
    {
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").FormulaA1 = "=1";
            ws.Cell("A3").FormulaA1 = "=1";
            ws.Cell("A4").FormulaA1 = "=1";
            ws.Cell("B1").FormulaA1 = "SUM(A:A)";
            ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        }

        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").SetFormulaA1("=1");
            ws.Cell("A3").SetFormulaA1("=1");
            ws.Cell("A4").SetFormulaA1("=1");
            ws.Cell("B1").SetFormulaA1("SUM(A:A)");
            ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        }

        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").FormulaR1C1 = "=1";
            ws.Cell("A3").FormulaR1C1 = "=1";
            ws.Cell("A4").FormulaR1C1 = "=1";
            ws.Cell("B1").FormulaR1C1 = "SUM(A:A)";
            ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        }

        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").SetFormulaR1C1("=1");
            ws.Cell("A3").SetFormulaR1C1("=1");
            ws.Cell("A4").SetFormulaR1C1("=1");
            ws.Cell("B1").SetFormulaR1C1("SUM(A:A)");
            ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        }
    }

    [Test]
    public void MergeSingleCellRangeDoesNothing()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.Range(1, 1, 1, 1);

        range.Merge();

        ClassicAssert.IsFalse(range.IsMerged());
        ClassicAssert.AreEqual(0, ws.MergedRanges.Count);
    }
}
