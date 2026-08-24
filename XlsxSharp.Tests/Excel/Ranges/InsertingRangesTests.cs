using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Tests.Excel.Ranges;

public class InsertingRangesTests
{
    [Test]
    public void InsertingColumnsPreservesFormatting()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");
        IXLColumn column1 = ws.Column(1);
        column1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
        column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
        IXLColumn column2 = ws.Column(2);
        column2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
        column2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

        column1.InsertColumnsAfter(1);
        column1.InsertColumnsBefore(1);
        column2.InsertColumnsBefore(1);

        ClassicAssert.AreEqual(
            ws.Style.Fill.BackgroundColor,
            ws.Column(1).Style.Fill.BackgroundColor
        );
        ClassicAssert.AreEqual(XLColor.FrenchLilac, ws.Column(2).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.FrenchLilac, ws.Column(3).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.FrenchLilac, ws.Column(4).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.Xanadu, ws.Column(5).Style.Fill.BackgroundColor);

        ClassicAssert.AreEqual(
            ws.Style.Fill.BackgroundColor,
            ws.Cell(2, 1).Style.Fill.BackgroundColor
        );
        ClassicAssert.AreEqual(XLColor.Fulvous, ws.Cell(2, 2).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.Fulvous, ws.Cell(2, 3).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.Fulvous, ws.Cell(2, 4).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.MacaroniAndCheese, ws.Cell(2, 5).Style.Fill.BackgroundColor);
    }

    [Test]
    public void InsertingRowsAbove()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");

        ws.Cell("B3").SetValue("X").CellBelow().SetValue("B");

        IXLRangeRow r = ws.Range("B4").InsertRowsAbove(1).First();
        r.Cell(1).SetValue("A");

        ClassicAssert.AreEqual("X", ws.Cell("B3").GetText());
        ClassicAssert.AreEqual("A", ws.Cell("B4").GetText());
        ClassicAssert.AreEqual("B", ws.Cell("B5").GetText());
    }

    [Test]
    public void InsertingRowsPreservesFormatting()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");
        IXLRow row1 = ws.Row(1);
        row1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
        row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
        IXLRow row2 = ws.Row(2);
        row2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
        row2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

        row1.InsertRowsBelow(1);
        row1.InsertRowsAbove(1);
        row2.InsertRowsAbove(1);

        ClassicAssert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.FrenchLilac, ws.Row(2).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.FrenchLilac, ws.Row(3).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.FrenchLilac, ws.Row(4).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.Xanadu, ws.Row(5).Style.Fill.BackgroundColor);

        ClassicAssert.AreEqual(
            ws.Style.Fill.BackgroundColor,
            ws.Cell(1, 2).Style.Fill.BackgroundColor
        );
        ClassicAssert.AreEqual(XLColor.Fulvous, ws.Cell(2, 2).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.Fulvous, ws.Cell(3, 2).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.Fulvous, ws.Cell(4, 2).Style.Fill.BackgroundColor);
        ClassicAssert.AreEqual(XLColor.MacaroniAndCheese, ws.Cell(5, 2).Style.Fill.BackgroundColor);
    }

    [Test]
    public void InsertingRowsPreservesComments()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").SetValue("Insert Below");
        ws.Cell("A2").SetValue("Already existing cell");
        ws.Cell("A3").SetValue("Cell with comment").GetComment().AddText("Comment here");

        ws.Row(1).InsertRowsBelow(2);
        ClassicAssert.AreEqual("Comment here", ws.Cell("A5").GetComment().Text);
    }

    [Test]
    public void InsertingColumnsPreservesComments()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").SetValue("Insert to the right");
        ws.Cell("B1").SetValue("Already existing cell");
        ws.Cell("C1").SetValue("Cell with comment").GetComment().AddText("Comment here");

        ws.Column(1).InsertColumnsAfter(2);
        ClassicAssert.AreEqual("Comment here", ws.Cell("E1").GetComment().Text);
    }

    [Test]
    [Arguments("C4:F7", "C4:F7", 2, "E4:H7")] // Coincide, shift right
    [Arguments("C4:F7", "C4:F7", -2, "C4:D7")] // Coincide, shift left
    [Arguments("D5:E6", "C4:F7", 2, "F5:G6")] // Inside, shift right
    [Arguments("D5:E6", "C4:F7", -2, "C5:C6")] // Inside, shift left
    [Arguments("B4:G7", "C4:F7", 2, "B4:I7")] // Includes, shift right
    [Arguments("B4:G7", "C4:F7", -2, "B4:E7")] // Includes, shift left
    [Arguments("B4:E7", "C4:F7", 2, "B4:G7")] // Intersects at left, shift right
    [Arguments("B4:E7", "C4:F7", -2, "B4:C7")] // Intersects at left, shift left
    [Arguments("D4:G7", "C4:F7", 2, "F4:I7")] // Intersects at right, shift right
    [Arguments("D4:G7", "C4:F7", -2, "C4:E7")] // Intersects at right, shift left
    [Arguments("A5:B6", "C4:F7", 2, "A5:B6")] // No intersection, at left, shift right
    [Arguments("A5:B6", "C4:F7", -1, "A5:B6")] // No intersection, at left, shift left
    [Arguments("H5:I6", "C4:F7", 2, "J5:K6")] // No intersection, at right, shift right
    [Arguments("H5:I6", "C4:F7", -2, "F5:G6")] // No intersection, at right, shift left
    [Arguments("C8:F11", "C4:F7", 2, "C8:F11")] // Different rows
    [Arguments("B1:B8", "A1:C4", 1, "B1:B8")] // More rows, shift right
    [Arguments("B1:B8", "A1:C4", -1, "B1:B8")] // More rows, shift left
    public void ShiftColumnsValid(
        string thisRangeAddress,
        string shiftedRangeAddress,
        int shiftedColumns,
        string expectedRange
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            XLRange? thisRange = ws.Range(thisRangeAddress) as XLRange;
            XLRange? shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

            thisRange.WorksheetRangeShiftedColumns(shiftedRange, shiftedColumns);

            ClassicAssert.IsTrue(thisRange.RangeAddress.IsValid);
            ClassicAssert.AreEqual(expectedRange, thisRange.RangeAddress.ToString());
        }
    }

    [Test]
    [Arguments("B1:B4", "A1:C4", -2)] // Shift left too much
    public void ShiftColumnsInvalid(
        string thisRangeAddress,
        string shiftedRangeAddress,
        int shiftedColumns
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            XLRange? thisRange = ws.Range(thisRangeAddress) as XLRange;
            XLRange? shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

            thisRange.WorksheetRangeShiftedColumns(shiftedRange, shiftedColumns);

            ClassicAssert.IsFalse(thisRange.RangeAddress.IsValid);
        }
    }

    [Test]
    [Arguments("C4:F7", "C4:F7", 2, "C6:F9")] // Coincide, shift down
    [Arguments("C4:F7", "C4:F7", -2, "C4:F5")] // Coincide, shift up
    [Arguments("D5:E6", "C4:F7", 2, "D7:E8")] // Inside, shift down
    [Arguments("D5:E6", "C4:F7", -2, "D4:E4")] // Inside, shift up
    [Arguments("C3:F8", "C4:F7", 2, "C3:F10")] // Includes, shift down
    [Arguments("C3:F8", "C4:F7", -2, "C3:F6")] // Includes, shift up
    [Arguments("C3:F6", "C4:F7", 2, "C3:F8")] // Intersects at top, shift down
    [Arguments("C2:F6", "C4:F7", -3, "C2:F3")] // Intersects at top, shift up to the sheet boundary
    [Arguments("C3:F6", "C4:F7", -2, "C3:F4")] // Intersects at top, shift up
    [Arguments("C5:F8", "C4:F7", 2, "C7:F10")] // Intersects at bottom, shift down
    [Arguments("C5:F8", "C4:F7", -2, "C4:F6")] // Intersects at bottom, shift up
    [Arguments("C1:F3", "C4:F7", 2, "C1:F3")] // No intersection, at top, shift down
    [Arguments("C1:F3", "C4:F7", -2, "C1:F3")] // No intersection, at top, shift up
    [Arguments("C8:F10", "C4:F7", 2, "C10:F12")] // No intersection, at bottom, shift down
    [Arguments("C8:F10", "C4:F7", -2, "C6:F8")] // No intersection, at bottom, shift up
    [Arguments("G4:J7", "C4:F7", 2, "G4:J7")] // Different columns
    [Arguments("A2:D2", "A1:C4", 1, "A2:D2")] // More columns, shift down
    [Arguments("A2:D2", "A1:C4", -1, "A2:D2")] // More columns, shift up
    public void ShiftRowsValid(
        string thisRangeAddress,
        string shiftedRangeAddress,
        int shiftedRows,
        string expectedRange
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            XLRange? thisRange = ws.Range(thisRangeAddress) as XLRange;
            XLRange? shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

            thisRange.WorksheetRangeShiftedRows(shiftedRange, shiftedRows);

            ClassicAssert.IsTrue(thisRange.RangeAddress.IsValid);
            ClassicAssert.AreEqual(expectedRange, thisRange.RangeAddress.ToString());
        }
    }

    [Test]
    [Arguments("A2:C2", "A1:C4", -2)] // Shift up too much
    public void ShiftRowsInvalid(
        string thisRangeAddress,
        string shiftedRangeAddress,
        int shiftedRows
    )
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            XLRange? thisRange = ws.Range(thisRangeAddress) as XLRange;
            XLRange? shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

            thisRange.WorksheetRangeShiftedRows(shiftedRange, shiftedRows);

            ClassicAssert.IsFalse(thisRange.RangeAddress.IsValid);
        }
    }

    [Test]
    public void InsertZeroColumnsFails()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        IXLRange range = ws.FirstCell().AsRange();
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertColumnsAfter(0)
        );
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertColumnsBefore(0)
        );
    }

    [Test]
    public void InsertNegativeNumberOfColumnsFails()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        IXLRange range = ws.FirstCell().AsRange();
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertColumnsAfter(-1)
        );
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertColumnsBefore(-1)
        );
    }

    [Test]
    public void InsertTooLargeNumberOfColumnsFails()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        IXLRange range = ws.FirstCell().AsRange();
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertColumnsAfter(16385)
        );
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertColumnsBefore(16385)
        );
    }

    [Test]
    public void InsertZeroRowsFails()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        IXLRange range = ws.FirstCell().AsRange();
        ClassicAssert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsAbove(0));
        ClassicAssert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsBelow(0));
    }

    [Test]
    public void InsertNegativeNumberOfRowsFails()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        IXLRange range = ws.FirstCell().AsRange();
        ClassicAssert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsAbove(-1));
        ClassicAssert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsBelow(-1));
    }

    [Test]
    public void InsertTooLargeNumberOrRowsFails()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        IXLRange range = ws.FirstCell().AsRange();
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertRowsAbove(1048577)
        );
        ClassicAssert.Throws(
            typeof(ArgumentOutOfRangeException),
            () => range.InsertRowsBelow(1048577)
        );
    }

    [Test]
    public void MergedRangesConsistencyWhenInsertingRows()
    {
        // https://github.com/XlsxSharp/XlsxSharp/issues/1013
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            //create merged row
            ws.Cell("A1").Value = "Merged Row(1) of Range (A1:F1)";
            ws.Range("A1:F1").Row(1).Merge();

            IXLRow row = ws.FirstRow();

            // Add some lines and copy format & merging
            for (int r = 1; r <= 10; r++)
            {
                row.InsertRowsBelow(1); // insert a row below row 1, as a row 2
                row.CopyTo(row.RowBelow()); // copy format and merging from row 1 to row 2

                var duplicates = ws
                    .MergedRanges.GroupBy(s => s.ToString())
                    .Where(g => g.Count() > 1)
                    .Select(y => new { Element = y.Key, Counter = y.Count() })
                    .ToList();

                ClassicAssert.AreEqual(0, duplicates.Count);
            }
        }
    }
}
