using System;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Examples.Misc;

public class CopyingRowsAndColumns : IXLExample
{
    public void Create(String filePath)
    {
        XLWorkbook workbook = new();

        IXLWorksheet originalSheet = workbook.Worksheets.Add("original");

        originalSheet.Cell("A2").SetValue("test value");
        originalSheet.Range("A2:E2").Merge();

        originalSheet.Cell("F1").SetValue("test value").Style.Alignment.SetTopToBottom();
        originalSheet.Range("F1:F6").Merge();

        IXLWorksheet fromRow = workbook.Worksheets.Add("From a Row");
        fromRow.Cell(1, 1).SetValue("Row to Row:");
        originalSheet.Row(2).CopyTo(fromRow.Row(2));
        fromRow.Cell(3, 1).SetValue("Row to Range:");
        originalSheet.Row(2).CopyTo(fromRow.Row(4).AsRange());
        fromRow.Cell(5, 1).SetValue("Row to Cell:");
        originalSheet.Row(2).CopyTo(fromRow.Row(6).FirstCell());

        IXLWorksheet fromRange = workbook.Worksheets.Add("From a Range");
        fromRange.Cell(1, 1).SetValue("Range to Row:");
        originalSheet.Row(2).AsRange().CopyTo(fromRange.Row(2));
        fromRange.Cell(3, 1).SetValue("Range to Range:");
        originalSheet.Row(2).AsRange().CopyTo(fromRange.Row(4).AsRange());
        fromRange.Cell(5, 1).SetValue("Range to Cell:");
        originalSheet.Row(2).AsRange().CopyTo(fromRange.Row(6).FirstCell());

        CopyRowAsRange(originalSheet, 2, fromRange, 8);

        IXLWorksheet fromColumn = workbook.Worksheets.Add("From a Column to Column");
        fromColumn.Cell(1, 1).SetValue("Column to Column:").Style.Alignment.SetTopToBottom();
        originalSheet.Column("F").CopyTo(fromColumn.Column(2));
        fromColumn.Cell(1, 3).SetValue("Column to Range:").Style.Alignment.SetTopToBottom();
        originalSheet.Column("F").CopyTo(fromColumn.Column(4).AsRange());
        fromColumn.Cell(1, 5).SetValue("Column to Cell:").Style.Alignment.SetTopToBottom();
        originalSheet.Column("F").CopyTo(fromColumn.Column(6).FirstCell());

        IXLWorksheet fromRangeToColumn = workbook.Worksheets.Add("From a Range to Column");
        fromRangeToColumn.Cell(1, 1).SetValue("Range to Column:").Style.Alignment.SetTopToBottom();
        originalSheet.Column("F").AsRange().CopyTo(fromRangeToColumn.Column(2));
        fromRangeToColumn.Cell(1, 3).SetValue("Range to Range:").Style.Alignment.SetTopToBottom();
        originalSheet.Column("F").AsRange().CopyTo(fromRangeToColumn.Column(4).AsRange());
        fromRangeToColumn.Cell(1, 5).SetValue("Range to Cell:").Style.Alignment.SetTopToBottom();
        originalSheet.Column("F").AsRange().CopyTo(fromRangeToColumn.Column(6).FirstCell());

        workbook.SaveAs(filePath);
    }

    private static void CopyRowAsRange(
        IXLWorksheet originalSheet,
        int originalRowNumber,
        IXLWorksheet destSheet,
        int destRowNumber
    )
    {
        IXLRow destinationRow = destSheet.Row(destRowNumber);
        destinationRow.Clear();

        IXLRow originalRow = originalSheet.Row(originalRowNumber);
        int columnNumber = originalRow.LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;

        IXLRange originalRange = originalSheet.Range(
            originalRowNumber,
            1,
            originalRowNumber,
            columnNumber
        );
        IXLRange destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
        originalRange.CopyTo(destRange);
    }
}
