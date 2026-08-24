using System.Collections;
using System.Data;
using XlsxSharp.Attributes;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.Tables;

public class AppendingAndReplacingTableDataTests
{
    public class TestObjectWithoutAttributes
    {
        public string Column1 { get; set; }
        public string Column2 { get; set; }
    }

    public class Person
    {
        public int Age { get; set; }

        [XLColumn(Header = "Last name", Order = 2)]
        public string LastName { get; set; }

        [XLColumn(Header = "First name", Order = 1)]
        public string FirstName { get; set; }

        [XLColumn(Header = "Full name", Order = 0)]
        public string FullName => string.Concat(this.FirstName, " ", this.LastName);

        [XLColumn(Order = 3)]
        public DateTime DateOfBirth { get; set; }

        [XLColumn(Header = "Is active", Order = 4)]
        public bool IsActive;
    }

    private static XLWorkbook PrepareWorkbook()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Tables");

        Person[] data =
        [
            new()
            {
                FirstName = "Francois",
                LastName = "Botha",
                Age = 39,
                DateOfBirth = new DateTime(1980, 1, 1),
                IsActive = true,
            },
            new()
            {
                FirstName = "Leon",
                LastName = "Oosthuizen",
                Age = 40,
                DateOfBirth = new DateTime(1979, 1, 1),
                IsActive = false,
            },
            new()
            {
                FirstName = "Rian",
                LastName = "Prinsloo",
                Age = 41,
                DateOfBirth = new DateTime(1978, 1, 1),
                IsActive = false,
            },
        ];

        ws.FirstCell().CellRight().CellBelow().InsertTable(data);

        ws.Columns().AdjustToContents();

        return wb;
    }

    private XLWorkbook PrepareWorkbookWithAdditionalColumns()
    {
        XLWorkbook wb = PrepareWorkbook();
        IXLWorksheet ws = wb.Worksheets.First();

        IXLTable table = ws.Tables.First();
        table
            .HeadersRow()
            .LastCell()
            .CellRight()
            .InsertData(
                new[] { "CumulativeAge", "NameLength", "IsOld", "HardCodedValue" },
                transpose: true
            );

        table.Resize(ws.Range(table.FirstCell(), table.LastCell().CellRight(4)));

        table
            .Field("CumulativeAge")
            .DataCells.ForEach(c => c.FormulaA1 = $"SUM($G$3:G{c.WorksheetRow().RowNumber()})");
        table
            .Field("NameLength")
            .DataCells.ForEach(c => c.FormulaA1 = $"LEN(B{c.WorksheetRow().RowNumber()})");
        table
            .Field("IsOld")
            .DataCells.ForEach(c => c.FormulaA1 = $"=G{c.WorksheetRow().RowNumber()}>=40");
        table.Field("HardCodedValue").DataCells.Value = "40 is not old!";

        return wb;
    }

    private static Person[] NewData =>
        [
            new()
            {
                FirstName = "Michelle",
                LastName = "de Beer",
                Age = 35,
                DateOfBirth = new DateTime(1983, 1, 1),
                IsActive = false,
            },
            new()
            {
                FirstName = "Marichen",
                LastName = "van der Gryp",
                Age = 30,
                DateOfBirth = new DateTime(1990, 1, 1),
                IsActive = true,
            },
        ];

    [Test]
    public void AddingEmptyEnumerables()
    {
        using (XLWorkbook wb = PrepareWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            IXLTable table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = null;
            ClassicAssert.AreEqual(null, table.AppendData(personEnumerable));

            personEnumerable = new Person[] { };
            ClassicAssert.AreEqual(null, table.AppendData(personEnumerable));

            IEnumerable enumerable = null;
            ClassicAssert.AreEqual(null, table.AppendData(enumerable));

            enumerable = new Person[] { };
            ClassicAssert.AreEqual(null, table.AppendData(enumerable));
        }
    }

    [Test]
    public void ReplaceWithEmptyEnumerables()
    {
        using (XLWorkbook wb = PrepareWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            IXLTable table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = null;
            ClassicAssert.Throws<InvalidOperationException>(() =>
                table.ReplaceData(personEnumerable)
            );

            personEnumerable = new Person[] { };
            ClassicAssert.Throws<InvalidOperationException>(() =>
                table.ReplaceData(personEnumerable)
            );

            IEnumerable enumerable = null;
            ClassicAssert.Throws<InvalidOperationException>(() => table.ReplaceData(enumerable));

            enumerable = new Person[] { };
            ClassicAssert.Throws<InvalidOperationException>(() => table.ReplaceData(enumerable));
        }
    }

    [Test]
    public void CanAppendTypedEnumerable()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;
                IXLRange addedRange = table.AppendData(personEnumerable);

                ClassicAssert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(5, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanAppendToTableWithTotalsRow()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();
                table.SetShowTotalsRow(true);
                table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                IEnumerable<Person> personEnumerable = NewData;
                IXLRange addedRange = table.AppendData(personEnumerable);

                ClassicAssert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(5, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanAppendTypedEnumerableAndPushDownCellsBelowTable()
    {
        using (MemoryStream ms = new())
        {
            string value = "Some value that will be overwritten";
            IXLAddress address;
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IXLCell cell = table.LastRow().FirstCell().CellRight(2).CellBelow(1);
                address = cell.Address;
                cell.Value = value;

                IEnumerable<Person> personEnumerable = NewData;
                IXLRange addedRange = table.AppendData(personEnumerable);

                ClassicAssert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IXLCell cell = ws.Cell(address);
                ClassicAssert.AreEqual("de Beer", cell.Value);
                ClassicAssert.AreEqual(5, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());

                ClassicAssert.AreEqual(value, cell.CellBelow(NewData.Length).Value);
            }
        }
    }

    [Test]
    public void CanAppendUntypedEnumerable()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                ArrayList list = new();
                list.AddRange(NewData);

                IXLRange addedRange = table.AppendData(list);

                ClassicAssert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(5, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanAppendDataTable()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;

                IXLWorksheet ws2 = wb.AddWorksheet("temp");
                DataTable dataTable = ws2.FirstCell()
                    .InsertTable(personEnumerable)
                    .AsNativeDataTable();

                IXLRange addedRange = table.AppendData(dataTable);

                ClassicAssert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(5, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanReplaceWithTypedEnumerable()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;
                IXLRange replacedRange = table.ReplaceData(personEnumerable);

                ClassicAssert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(2, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanReplaceWithUntypedEnumerable()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                ArrayList list = new();
                list.AddRange(NewData);

                IXLRange replacedRange = table.ReplaceData(list);

                ClassicAssert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(2, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanReplaceWithDataTable()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;

                IXLWorksheet ws2 = wb.AddWorksheet("temp");
                DataTable dataTable = ws2.FirstCell()
                    .InsertTable(personEnumerable)
                    .AsNativeDataTable();

                IXLRange replacedRange = table.ReplaceData(dataTable);

                ClassicAssert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(2, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanReplaceToTableWithTablesRow1()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();
                table.SetShowTotalsRow(true);
                table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                // Will cause table to overflow
                IEnumerable<Person> personEnumerable = NewData.Union(NewData).Union(NewData);
                IXLRange replacedRange = table.ReplaceData(personEnumerable);

                ClassicAssert.AreEqual("B3:G8", replacedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(6, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanReplaceToTableWithTablesRow2()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();
                table.SetShowTotalsRow(true);
                table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                // Will cause table to shrink
                IEnumerable<Person> personEnumerable = NewData.Take(1);
                IXLRange replacedRange = table.ReplaceData(personEnumerable);

                ClassicAssert.AreEqual("B3:G3", replacedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(1, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanReplaceWithUntypedEnumerableAndPropagateExtraColumns()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = this.PrepareWorkbookWithAdditionalColumns())
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLTable table = ws.Tables.First();

                ArrayList list = new();
                list.AddRange(NewData);
                list.AddRange(NewData);

                IXLRange replacedRange = table.ReplaceData(list, propagateExtraColumns: true);

                ClassicAssert.AreEqual("B3:G6", replacedRange.RangeAddress.ToString());

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(4, table.DataRange.RowCount());
                ClassicAssert.AreEqual(10, table.DataRange.ColumnCount());

                ClassicAssert.AreEqual("SUM($G$3:G5)", table.Worksheet.Cell("H5").FormulaA1);
                ClassicAssert.AreEqual("SUM($G$3:G6)", table.Worksheet.Cell("H6").FormulaA1);
                ClassicAssert.AreEqual(100, table.Worksheet.Cell("H5").Value);
                ClassicAssert.AreEqual(130, table.Worksheet.Cell("H6").Value);

                ClassicAssert.AreEqual("LEN(B5)", table.Worksheet.Cell("I5").FormulaA1);
                ClassicAssert.AreEqual("LEN(B6)", table.Worksheet.Cell("I6").FormulaA1);
                ClassicAssert.AreEqual(16, table.Worksheet.Cell("I5").Value);
                ClassicAssert.AreEqual(21, table.Worksheet.Cell("I6").Value);

                ClassicAssert.AreEqual("G5>=40", table.Worksheet.Cell("J5").FormulaA1);
                ClassicAssert.AreEqual("G6>=40", table.Worksheet.Cell("J6").FormulaA1);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J5").Value);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J6").Value);

                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K5").Value);
                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K6").Value);
            }
        }
    }

    [Test]
    public void CanReplaceWithTypedEnumerableAndPropagateExtraColumns()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = this.PrepareWorkbookWithAdditionalColumns())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData.Concat(NewData).OrderBy(p => p.Age);
                IXLRange replacedRange = table.ReplaceData(
                    personEnumerable,
                    propagateExtraColumns: true
                );

                ClassicAssert.AreEqual("B3:G6", replacedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(4, table.DataRange.RowCount());
                ClassicAssert.AreEqual(10, table.DataRange.ColumnCount());

                ClassicAssert.AreEqual("SUM($G$3:G5)", table.Worksheet.Cell("H5").FormulaA1);
                ClassicAssert.AreEqual("SUM($G$3:G6)", table.Worksheet.Cell("H6").FormulaA1);
                ClassicAssert.AreEqual(95, table.Worksheet.Cell("H5").Value);
                ClassicAssert.AreEqual(130, table.Worksheet.Cell("H6").Value);

                ClassicAssert.AreEqual("LEN(B5)", table.Worksheet.Cell("I5").FormulaA1);
                ClassicAssert.AreEqual("LEN(B6)", table.Worksheet.Cell("I6").FormulaA1);
                ClassicAssert.AreEqual(16, table.Worksheet.Cell("I5").Value);
                ClassicAssert.AreEqual(16, table.Worksheet.Cell("I6").Value);

                ClassicAssert.AreEqual("G5>=40", table.Worksheet.Cell("J5").FormulaA1);
                ClassicAssert.AreEqual("G6>=40", table.Worksheet.Cell("J6").FormulaA1);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J5").Value);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J6").Value);

                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K5").Value);
                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K6").Value);
            }
        }
    }

    [Test]
    [Arguments("ListOfPeople[Age]")] // Defined name formula without a A1 reference
    [Arguments("ListOfPeople!A1")] // Defined name formula with an A1 reference
    public void CanReplaceTableDataWhenWorksheetHasDefinedNames(string nameFormula)
    {
        // When table data are replaced, the size of a table is modified. That
        // means rows below it are shifted up/down and defined names should be
        // adjusted.
        // TODO: add assert for name shift when formulas are properly shifted. Originally, it threw even on defined name with A1 reference
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = PrepareWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                ws.DefinedNames.Add("ListOfPeople_Age", nameFormula);

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;
                IXLRange replacedRange = table.ReplaceData(personEnumerable);

                ClassicAssert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(2, table.DataRange.RowCount());
                ClassicAssert.AreEqual(6, table.DataRange.ColumnCount());
            }
        }
    }

    [Test]
    public void CanAppendWithUntypedEnumerableAndPropagateExtraColumns()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = this.PrepareWorkbookWithAdditionalColumns())
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLTable table = ws.Tables.First();

                ArrayList list = new();
                list.AddRange(NewData);
                list.AddRange(NewData);

                IXLRange appendedRange = table.AppendData(list, propagateExtraColumns: true);

                ClassicAssert.AreEqual("B6:G9", appendedRange.RangeAddress.ToString());

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(7, table.DataRange.RowCount());
                ClassicAssert.AreEqual(10, table.DataRange.ColumnCount());

                ClassicAssert.AreEqual("SUM($G$3:G8)", table.Worksheet.Cell("H8").FormulaA1);
                ClassicAssert.AreEqual("SUM($G$3:G9)", table.Worksheet.Cell("H9").FormulaA1);
                ClassicAssert.AreEqual(220, table.Worksheet.Cell("H8").Value);
                ClassicAssert.AreEqual(250, table.Worksheet.Cell("H9").Value);

                ClassicAssert.AreEqual("LEN(B8)", table.Worksheet.Cell("I8").FormulaA1);
                ClassicAssert.AreEqual("LEN(B9)", table.Worksheet.Cell("I9").FormulaA1);
                ClassicAssert.AreEqual(16, table.Worksheet.Cell("I8").Value);
                ClassicAssert.AreEqual(21, table.Worksheet.Cell("I9").Value);

                ClassicAssert.AreEqual("G8>=40", table.Worksheet.Cell("J8").FormulaA1);
                ClassicAssert.AreEqual("G9>=40", table.Worksheet.Cell("J9").FormulaA1);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J8").Value);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J9").Value);

                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K8").Value);
                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K9").Value);
            }
        }
    }

    [Test]
    public void CanAppendTypedEnumerableAndPropagateExtraColumns()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = this.PrepareWorkbookWithAdditionalColumns())
            {
                IXLWorksheet ws = wb.Worksheets.First();

                IXLTable table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData
                    .Concat(NewData)
                    .Concat(NewData)
                    .OrderBy(p => p.FirstName);

                IXLRange addedRange = table.AppendData(personEnumerable);

                ClassicAssert.AreEqual("B6:G11", addedRange.RangeAddress.ToString());
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLTable table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                ClassicAssert.AreEqual(9, table.DataRange.RowCount());
                ClassicAssert.AreEqual(10, table.DataRange.ColumnCount());

                ClassicAssert.AreEqual("SUM($G$3:G10)", table.Worksheet.Cell("H10").FormulaA1);
                ClassicAssert.AreEqual("SUM($G$3:G11)", table.Worksheet.Cell("H11").FormulaA1);
                ClassicAssert.AreEqual(280, table.Worksheet.Cell("H10").Value);
                ClassicAssert.AreEqual(315, table.Worksheet.Cell("H11").Value);

                ClassicAssert.AreEqual("LEN(B10)", table.Worksheet.Cell("I10").FormulaA1);
                ClassicAssert.AreEqual("LEN(B11)", table.Worksheet.Cell("I11").FormulaA1);
                ClassicAssert.AreEqual(16, table.Worksheet.Cell("I10").Value);
                ClassicAssert.AreEqual(16, table.Worksheet.Cell("I11").Value);

                ClassicAssert.AreEqual("G10>=40", table.Worksheet.Cell("J10").FormulaA1);
                ClassicAssert.AreEqual("G11>=40", table.Worksheet.Cell("J11").FormulaA1);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J10").Value);
                ClassicAssert.AreEqual(false, table.Worksheet.Cell("J11").Value);

                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K10").Value);
                ClassicAssert.AreEqual("40 is not old!", table.Worksheet.Cell("K11").Value);
            }
        }
    }
}
