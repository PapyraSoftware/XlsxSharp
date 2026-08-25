using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Benchmarks;

// Writes a RowCount x ColumnCount grid of alternating number/string cells straight through the
// OpenXML SDK. Shared by WriteBenchmarks (it's the "raw SDK" baseline) and ReadBenchmarks (it
// builds the read fixtures, so no library under test gets to benchmark its own output).
internal static class OpenXmlGrid
{
    public static byte[] Write(int rowCount, int columnCount)
    {
        using MemoryStream? stream = new();
        using (
            SpreadsheetDocument? document = SpreadsheetDocument.Create(
                stream,
                SpreadsheetDocumentType.Workbook
            )
        )
        {
            WorkbookPart? workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // A real .xlsx always carries a styles part with at least one default cell format
            // (styleId 0). Without it, some readers - XlsxSharp included - fault resolving the
            // default style while computing column widths, so the fixture has to include one to
            // stay representative of an actual Excel-produced file.
            WorkbookStylesPart? stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()) { Count = 1 },
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                )
                {
                    Count = 2,
                },
                new Borders(new Border()) { Count = 1 },
                new CellFormats(
                    new CellFormat
                    {
                        FontId = 0,
                        FillId = 0,
                        BorderId = 0,
                        NumberFormatId = 0,
                    }
                )
                {
                    Count = 1,
                }
            );
            stylesPart.Stylesheet.Save();

            WorksheetPart? worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData? sheetData = new();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            for (int r = 1; r <= rowCount; r++)
            {
                Row? row = new() { RowIndex = (uint)r };
                for (int c = 0; c < columnCount; c++)
                {
                    row.Append(
                        c % 2 == 0
                            ? new Cell
                            {
                                CellReference = CellReference(r, c),
                                DataType = CellValues.Number,
                                CellValue = new CellValue(
                                    (r * columnCount + c).ToString(CultureInfo.InvariantCulture)
                                ),
                            }
                            : new Cell
                            {
                                CellReference = CellReference(r, c),
                                DataType = CellValues.InlineString,
                                InlineString = new InlineString(new Text($"Row {r} Col {c}")),
                            }
                    );
                }

                sheetData.Append(row);
            }

            Sheets? sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(
                new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1",
                }
            );
            workbookPart.Workbook.Save();
        }

        return stream.ToArray();
    }

    public static string CellReference(int row, int zeroBasedColumn) =>
        $"{ColumnLetter(zeroBasedColumn)}{row}";

    private static string ColumnLetter(int zeroBasedColumn)
    {
        int dividend = zeroBasedColumn + 1;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = (char)('A' + modulo) + columnName;
            dividend = (dividend - modulo - 1) / 26;
        }

        return columnName;
    }
}
