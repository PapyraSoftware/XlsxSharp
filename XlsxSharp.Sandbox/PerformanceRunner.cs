using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using XlsxSharp.Excel;

namespace XlsxSharp.Sandbox;

internal class PerformanceRunner
{
    public static void TimeAction(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        Console.WriteLine("Action done in " + stopwatch.Elapsed);
    }

    private const int rowCount = 5000;

    public static void RunInsertTable()
    {
        List<OneRow> rows = [];

        for (int i = 0; i < rowCount; i++)
        {
            OneRow row = GenerateRow<OneRow>();
            rows.Add(row);
        }

        XLWorkbook workbook = new();
        IXLWorksheet worksheet = workbook.Worksheets.Add("Sheet 1");
        worksheet.Cell(1, 1).InsertTable(rows);

        CreateMergedCell(worksheet);

        worksheet.Columns().AdjustToContents();

        EmulateSave(workbook);
    }

    public static void OpenTestFile()
    {
        using (XLWorkbook wb = new("test.xlsx"))
        {
            wb.RecalculateAllFormulas();
            IXLWorksheet ws = wb.Worksheets.First();
            IXLCell cell = ws.FirstCellUsed();
            Console.WriteLine(cell.Value);
        }
    }

    private static void CreateMergedCell(IXLWorksheet worksheet)
    {
        worksheet.Cell(rowCount + 2, 1).Value = "Merged cell";
        IXLRange range = worksheet.Range(rowCount + 2, 1, rowCount + 2, 2);
        range.Row(1).Merge();
    }

    private static void EmulateSave(XLWorkbook workbook)
    {
        using (MemoryStream memoryStream = new())
        {
            workbook.SaveAs(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);
            Console.WriteLine("Total bytes = " + memoryStream.ToArray().Length);
        }
    }

    private static Random rnd = new();

    private static T GenerateRow<T>()
        where T : new()
    {
        T row = new();

        PropertyInfo[] rowProps = row.GetType().GetProperties();

        IEnumerable<PropertyInfo> strings = rowProps.Where(p => p.PropertyType == typeof(string));
        IEnumerable<PropertyInfo> decimals = rowProps.Where(p => p.PropertyType == typeof(decimal));
        IEnumerable<PropertyInfo> ints = rowProps.Where(p =>
            p.PropertyType == typeof(int) || p.PropertyType == typeof(int?)
        );
        IEnumerable<PropertyInfo> dates = rowProps.Where(p => p.PropertyType == typeof(DateTime?));
        IEnumerable<PropertyInfo> timeSpans = rowProps.Where(p =>
            p.PropertyType == typeof(TimeSpan?)
        );
        IEnumerable<PropertyInfo> booleans = rowProps.Where(p => p.PropertyType == typeof(bool));

        // Format strings
        StringBuilder tmpString = new();
        int tmpStringLength = rnd.Next(5, 50);
        for (int x = 0; x <= tmpStringLength; x++)
        {
            tmpString.Append((char)(rnd.Next(48, 120)));
        }
        foreach (PropertyInfo str in strings)
        {
            str.SetValue(row, tmpString.ToString(), null);
        }

        // Format decimals
        decimal tmpDec = (decimal)(rnd.Next(-10000, 100000) / (Math.Pow(10.0, rnd.Next(1, 4))));

        foreach (PropertyInfo dec in decimals)
        {
            dec.SetValue(row, tmpDec, null);
        }

        // Format ints
        int tmpInt = rnd.Next(-1000, 10000);

        foreach (PropertyInfo intValue in ints)
        {
            intValue.SetValue(row, tmpInt, null);
        }

        // Format dates
        DateTime tmpDate = new(2012, 1, 1, 1, 1, 1);
        tmpDate = tmpDate.AddSeconds(rnd.Next(-10000, 100000));
        foreach (PropertyInfo dt in dates)
        {
            dt.SetValue(row, tmpDate, null);
        }

        // Format timespans
        TimeSpan tmpTimespan = new(rnd.Next(1, 24), rnd.Next(1, 60), rnd.Next(1, 60));

        foreach (PropertyInfo ts in timeSpans)
        {
            ts.SetValue(row, tmpTimespan, null);
        }

        // Format booleans
        bool tmpBool = (rnd.Next(0, 2) > 0);
        foreach (PropertyInfo bl in booleans)
        {
            bl.SetValue(row, tmpBool, null);
        }

        return row;
    }

    public static void PerformHeavyCalculation()
    {
        int rows = 200;
        int columns = 200;
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            string lastColumnLetter = sheet.Column(columns).ColumnLetter();
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    if (i == 1)
                    {
                        sheet.Cell(i, j).FormulaA1 = string.Format("=ROUND({0}*SIN({0}),2)", j);
                    }
                    else
                    {
                        sheet.Cell(i, j).FormulaA1 = string.Format(
                            "=SUM({0}$1:{0}{1})/SUM($A{1}:${2}{1})",
                            sheet.Column(j).ColumnLetter(),
                            i - 1,
                            lastColumnLetter
                        ); // i.e. for K8 there will be =SUM(K$1:K7)/SUM($A7:$GR7)
                    }
                }
            }

            IXLCells cells = sheet.CellsUsed();
            double sum1 = cells.Sum(cell => (double)cell.Value);
            Console.WriteLine("Total sum: {0:N2}", sum1);
        }
    }
}
