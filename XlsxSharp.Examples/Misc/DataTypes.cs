using System;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Examples.Misc;

public class DataTypes : IXLExample
{
    public void Create(String filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Data Types");

        int co = 2;
        int ro = 1;

        ws.Cell(++ro, co).Value = "Plain Text:";
        ws.Cell(ro, co + 1).Value = "Hello World.";

        ws.Cell(++ro, co).Value = "Plain Date:";
        ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);

        ws.Cell(++ro, co).Value = "Plain DateTime:";
        ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);

        ws.Cell(++ro, co).Value = "Plain Boolean:";
        ws.Cell(ro, co + 1).Value = true;

        ws.Cell(++ro, co).Value = "Plain Number:";
        ws.Cell(ro, co + 1).Value = 123.45;

        ws.Cell(++ro, co).Value = "TimeSpan:";
        ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);

        ro++;

        ws.Cell(++ro, co).Value = "Large Double Number:";
        ws.Cell(ro, co + 1).Value = 9.999E307d;

        ro++;

        ws.Cell(++ro, co).Value = "Explicit Text:";
        ws.Cell(ro, co + 1).Value = "'Hello World.";

        ws.Cell(++ro, co).Value = "Date as Text:";
        ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2);

        ws.Cell(++ro, co).Value = "DateTime as Text:";
        ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22);

        ws.Cell(++ro, co).Value = "Boolean as Text:";
        ws.Cell(ro, co + 1).Value = "'TRUE";

        ws.Cell(++ro, co).Value = "Number as Text:";
        ws.Cell(ro, co + 1).Value = "'123.45";

        ws.Cell(++ro, co).Value = "Number with @ format:";
        ws.Cell(ro, co + 1).Style.NumberFormat.Format = "@";
        ws.Cell(ro, co + 1).Value = 123.45;

        ws.Cell(++ro, co).Value = "Format number with @:";
        ws.Cell(ro, co + 1).Value = 123.45;
        ws.Cell(ro, co + 1).Style.NumberFormat.Format = "@";

        ws.Cell(++ro, co).Value = "TimeSpan as Text:";
        ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22);

        // Using inline strings (few users will ever need to use this feature)
        //
        // By default all strings are stored as shared so one block of text
        // can be reference by multiple cells.
        // You can override this by setting the .ShareString property to false
        ws.Cell(++ro, co).Value = "Inline String:";
        IXLCell cell = ws.Cell(ro, co + 1);
        cell.Value = "Not Shared";
        cell.ShareString = false;

        ro++;

        ws.Cell(++ro, co).Value = "Error from literal:";
        ws.Cell(ro, co + 1).Value = XLError.IncompatibleValue;

        ws.Cell(++ro, co).Value = "Error from evaluation:";
        ws.Cell(ro, co + 1).FormulaA1 = "1/0";

        ws.Columns(2, 3).AdjustToContents();

        workbook.SaveAs(filePath);
    }
}
