using System.IO;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;

namespace XlsxSharp.Benchmarks;

// Writes a RowCount x ColumnCount grid of mixed string/number cells and returns the saved bytes.
// ClosedXML and XlsxSharp both expose a type named XLWorkbook, so those two calls are fully
// qualified instead of relying on a `using` for either namespace.
[MemoryDiagnoser]
public class WriteBenchmarks
{
    private const int ColumnCount = 10;

    [Params(1_000, 10_000)]
    public int RowCount { get; set; }

    [Benchmark(Baseline = true, Description = "OpenXML SDK")]
    public byte[] OpenXmlSdk() => OpenXmlGrid.Write(RowCount, ColumnCount);

    [Benchmark(Description = "ClosedXML")]
    public byte[] ClosedXml()
    {
        using XLWorkbook? workbook = new();
        IXLWorksheet? worksheet = workbook.Worksheets.Add("Sheet1");
        FillGrid(worksheet);

        using MemoryStream? stream = new();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    [Benchmark(Description = "XlsxSharp")]
    public byte[] XlsxSharp()
    {
        using Excel.XLWorkbook? workbook = new();
        Excel.IXLWorksheet? worksheet = workbook.Worksheets.Add("Sheet1");
        FillGrid(worksheet);

        using MemoryStream? stream = new();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    [Benchmark(Description = "EPPlus")]
    public byte[] EPPlus()
    {
        using ExcelPackage? package = new();
        ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Sheet1");

        for (int r = 1; r <= RowCount; r++)
        {
            for (int c = 1; c <= ColumnCount; c++)
            {
                if (c % 2 == 1)
                {
                    worksheet.Cells[r, c].Value = r * ColumnCount + c;
                }
                else
                {
                    worksheet.Cells[r, c].Value = $"Row {r} Col {c}";
                }
            }
        }

        return package.GetAsByteArray();
    }

    [Benchmark(Description = "NPOI")]
    public byte[] Npoi()
    {
        XSSFWorkbook? workbook = new();
        ISheet sheet = workbook.CreateSheet("Sheet1");

        for (int r = 0; r < RowCount; r++)
        {
            IRow row = sheet.CreateRow(r);
            for (int c = 0; c < ColumnCount; c++)
            {
                ICell cell = row.CreateCell(c);
                if (c % 2 == 0)
                {
                    cell.SetCellValue(r * ColumnCount + c);
                }
                else
                {
                    cell.SetCellValue($"Row {r} Col {c}");
                }
            }
        }

        using MemoryStream? stream = new();
        workbook.Write(stream, leaveOpen: true);
        return stream.ToArray();
    }

    // ClosedXML.Excel.IXLWorksheet and XlsxSharp.Excel.IXLWorksheet have the same shape but no
    // shared base type across the two libraries, so the fill logic is duplicated per overload
    // instead of factored into one generic helper.
    private void FillGrid(ClosedXML.Excel.IXLWorksheet worksheet)
    {
        for (int r = 1; r <= RowCount; r++)
        {
            for (int c = 1; c <= ColumnCount; c++)
            {
                if (c % 2 == 1)
                {
                    worksheet.Cell(r, c).Value = r * ColumnCount + c;
                }
                else
                {
                    worksheet.Cell(r, c).Value = $"Row {r} Col {c}";
                }
            }
        }
    }

    private void FillGrid(Excel.IXLWorksheet worksheet)
    {
        for (int r = 1; r <= RowCount; r++)
        {
            for (int c = 1; c <= ColumnCount; c++)
            {
                if (c % 2 == 1)
                {
                    worksheet.Cell(r, c).Value = r * ColumnCount + c;
                }
                else
                {
                    worksheet.Cell(r, c).Value = $"Row {r} Col {c}";
                }
            }
        }
    }
}
