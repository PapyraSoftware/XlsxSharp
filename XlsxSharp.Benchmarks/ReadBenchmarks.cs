using System.Globalization;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;

namespace XlsxSharp.Benchmarks;

// Reads a pre-generated RowCount x ColumnCount workbook and sums every numeric cell, so the whole
// sheet actually gets deserialized instead of being skipped by a lazy reader. The fixture file is
// built once per RowCount via OpenXmlGrid, so no library gets to benchmark its own output.
[MemoryDiagnoser]
public class ReadBenchmarks
{
    private const int ColumnCount = 10;
    private byte[] _fileBytes = [];

    [Params(1_000, 10_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        _fileBytes = OpenXmlGrid.Write(RowCount, ColumnCount);
    }

    [Benchmark(Baseline = true, Description = "OpenXML SDK")]
    public double OpenXmlSdk()
    {
        using var stream = new MemoryStream(_fileBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var worksheetPart = (WorksheetPart)
            workbookPart.GetPartById(workbookPart.Workbook.Descendants<Sheet>().First().Id!);

        double sum = 0;
        foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
        {
            if (cell.DataType?.Value == CellValues.Number && cell.CellValue is not null)
            {
                sum += double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);
            }
        }

        return sum;
    }

    [Benchmark(Description = "ClosedXML")]
    public double ClosedXml()
    {
        using var stream = new MemoryStream(_fileBytes);
        using var workbook = new ClosedXML.Excel.XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        double sum = 0;
        foreach (var cell in worksheet.CellsUsed())
        {
            if (cell.DataType == ClosedXML.Excel.XLDataType.Number)
            {
                sum += cell.GetDouble();
            }
        }

        return sum;
    }

    [Benchmark(Description = "XlsxSharp")]
    public double XlsxSharp()
    {
        using var stream = new MemoryStream(_fileBytes);
        using var workbook = new Excel.XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        double sum = 0;
        foreach (var cell in worksheet.CellsUsed())
        {
            if (cell.DataType == Excel.XLDataType.Number)
            {
                sum += cell.GetDouble();
            }
        }

        return sum;
    }

    [Benchmark(Description = "EPPlus")]
    public double EPPlus()
    {
        using var stream = new MemoryStream(_fileBytes);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets[0];

        double sum = 0;
        for (int r = 1; r <= RowCount; r++)
        {
            for (int c = 1; c <= ColumnCount; c++)
            {
                var cell = worksheet.Cells[r, c];
                if (cell.Value is double or int)
                {
                    sum += System.Convert.ToDouble(cell.Value, CultureInfo.InvariantCulture);
                }
            }
        }

        return sum;
    }

    [Benchmark(Description = "NPOI")]
    public double Npoi()
    {
        using var stream = new MemoryStream(_fileBytes);
        var workbook = new XSSFWorkbook(stream);
        ISheet sheet = workbook.GetSheetAt(0);

        double sum = 0;
        for (int r = 0; r <= sheet.LastRowNum; r++)
        {
            IRow? row = sheet.GetRow(r);
            if (row is null)
            {
                continue;
            }

            foreach (ICell cell in row)
            {
                if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    sum += cell.NumericCellValue;
                }
            }
        }

        return sum;
    }
}
