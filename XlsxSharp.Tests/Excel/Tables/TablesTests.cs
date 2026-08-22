using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Attributes;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.Exceptions;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.Tables;

[TestFixture]
public class TablesTests
{
    public class TestObjectWithoutAttributes
    {
        public String Column1 { get; set; }
        public String Column2 { get; set; }
    }

    public class TestObjectWithAttributes
    {
        public int UnOrderedColumn { get; set; }

        [XLColumn(Header = "SecondColumn", Order = 1)]
        public String Column1 { get; set; }

        [XLColumn(Header = "FirstColumn", Order = 0)]
        public String Column2 { get; set; }

        [XLColumn(Header = "SomeFieldNotProperty", Order = 2)]
        public int MyField;
    }

    [Test]
    public void CanSaveTableCreatedFromEmptyDataTable()
    {
        DataTable dt = new("sheet1");
        dt.Columns.Add("col1", typeof(string));
        dt.Columns.Add("col2", typeof(double));

        using (XLWorkbook wb = new())
        {
            wb.AddWorksheet(dt);

            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
            }
        }
    }

    [Test]
    public void PreventAddingOfEmptyDataTable()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            DataTable dt = new();
            IXLTable table = ws.FirstCell().InsertTable(dt);

            Assert.AreEqual(null, table);
        }
    }

    [Test]
    public void CanSaveTableCreatedFromSingleRow()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title");
            ws.Range("A1").CreateTable();

            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
            }
        }
    }

    [Test]
    public void CreatingATableFromHeadersPushCellsBelow()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title").CellBelow().SetValue("X");
            ws.Range("A1").CreateTable();

            Assert.AreEqual(Blank.Value, ws.Cell("A2").Value);
            Assert.AreEqual("X", ws.Cell("A3").GetText());
        }
    }

    [Test]
    public void InsertingColumnSetsHeader()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell()
                .SetValue("Categories")
                .CellBelow()
                .SetValue("A")
                .CellBelow()
                .SetValue("B")
                .CellBelow()
                .SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();
            table.InsertColumnsAfter(1);
            Assert.AreEqual("Column2", table.HeadersRow().LastCell().GetText());
        }
    }

    [Test]
    public void DataRangeReturnsNullIfEmpty()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell()
                .SetValue("Categories")
                .CellBelow()
                .SetValue("A")
                .CellBelow()
                .SetValue("B")
                .CellBelow()
                .SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();

            ws.Rows("2:4").Delete();

            Assert.IsNull(table.DataRange);
        }
    }

    [Test]
    public void SavingLoadingTableWithNewLineInHeader()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            string columnName = "Line1" + Environment.NewLine + "Line2";
            ws.FirstCell().SetValue(columnName).CellBelow().SetValue("A");
            ws.RangeUsed().CreateTable();
            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
                XLWorkbook wb2 = new(ms);
                IXLWorksheet ws2 = wb2.Worksheet(1);
                IXLTable table2 = ws2.Table(0);
                string fieldName = table2.Field(0).Name;
                Assert.AreEqual("Line1\nLine2", fieldName);
            }
        }
    }

    [Test]
    public void SavingLoadingTableWithNewLineInHeader2()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Test");

            DataTable dt = new();
            string columnName = "Line1" + Environment.NewLine + "Line2";
            dt.Columns.Add(columnName);

            DataRow dr = dt.NewRow();
            dr[columnName] = "some text";
            dt.Rows.Add(dr);
            ws.Cell(1, 1).InsertTable(dt);

            IXLTable table1 = ws.Table(0);
            string fieldName1 = table1.Field(0).Name;
            Assert.AreEqual(columnName, fieldName1);

            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
                XLWorkbook wb2 = new(ms);
                IXLWorksheet ws2 = wb2.Worksheet(1);
                IXLTable table2 = ws2.Table(0);
                string fieldName2 = table2.Field(0).Name;
                Assert.AreEqual("Line1\nLine2", fieldName2);
            }
        }
    }

    [Test]
    public void TableCreatedFromEmptyDataTable()
    {
        DataTable dt = new("sheet1");
        dt.Columns.Add("col1", typeof(string));
        dt.Columns.Add("col2", typeof(double));

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(dt);
            Assert.AreEqual(2, ws.Tables.First().ColumnCount());
        }
    }

    [Test]
    public void TableCreatedFromEmptyListOfInt()
    {
        List<int> l = [];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(1, ws.Tables.First().ColumnCount());
        }
    }

    [Test]
    public void TableCreatedFromEmptyListOfObject()
    {
        List<TestObjectWithoutAttributes> l = [];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(2, ws.Tables.First().ColumnCount());
        }
    }

    [Test]
    public void TableCreatedFromListOfObjectWithPropertyAttributes()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(4, ws.Tables.First().ColumnCount());
            Assert.AreEqual("FirstColumn", ws.FirstCell().Value);
            Assert.AreEqual("SecondColumn", ws.FirstCell().CellRight().Value);
            Assert.AreEqual("SomeFieldNotProperty", ws.FirstCell().CellRight().CellRight().Value);
            Assert.AreEqual(
                "UnOrderedColumn",
                ws.FirstCell().CellRight().CellRight().CellRight().Value
            );
        }
    }

    [Test]
    public void EmptyTableCreatedFromListOfObjectWithPropertyAttributes()
    {
        List<TestObjectWithAttributes> l = [];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(4, ws.Tables.First().ColumnCount());
            Assert.AreEqual("FirstColumn", ws.FirstCell().Value);
            Assert.AreEqual("SecondColumn", ws.FirstCell().CellRight().Value);
            Assert.AreEqual("SomeFieldNotProperty", ws.FirstCell().CellRight().CellRight().Value);
            Assert.AreEqual(
                "UnOrderedColumn",
                ws.FirstCell().CellRight().CellRight().CellRight().Value
            );
        }
    }

    [Test]
    public void TableInsertAboveFromData()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow().Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 3;
            row = table.DataRange.InsertRowsAbove(1).First();
            row.Field("Value").Value = 2;
            row = table.DataRange.InsertRowsAbove(1).First();
            row.Field("Value").Value = 1;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }
    }

    [Test]
    public void TableInsertAboveFromRows()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow().Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 3;
            row = row.InsertRowsAbove(1).First();
            row.Field("Value").Value = 2;
            row = row.InsertRowsAbove(1).First();
            row.Field("Value").Value = 1;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }
    }

    [Test]
    public void TableInsertBelowFromData()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow().Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 1;
            row = table.DataRange.InsertRowsBelow(1).First();
            row.Field("Value").Value = 2;
            row = table.DataRange.InsertRowsBelow(1).First();
            row.Field("Value").Value = 3;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }
    }

    [Test]
    public void TableInsertBelowFromRows()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow().Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 1;
            row = row.InsertRowsBelow(1).First();
            row.Field("Value").Value = 2;
            row = row.InsertRowsBelow(1).First();
            row.Field("Value").Value = 3;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }
    }

    [Test]
    public void TableShowHeader()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell()
                .SetValue("Categories")
                .CellBelow()
                .SetValue("A")
                .CellBelow()
                .SetValue("B")
                .CellBelow()
                .SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();

            Assert.AreEqual("Categories", table.Fields.First().Name);

            table.SetShowHeaderRow(false);

            Assert.AreEqual("Categories", table.Fields.First().Name);

            Assert.IsTrue(ws.Cell(1, 1).IsEmpty(XLCellsUsedOptions.All));
            Assert.AreEqual(null, table.HeadersRow());
            Assert.AreEqual("A", table.DataRange.FirstRow().Field("Categories").GetText());
            Assert.AreEqual("C", table.DataRange.LastRow().Field("Categories").GetText());
            Assert.AreEqual("A", table.DataRange.FirstCell().GetText());
            Assert.AreEqual("C", table.DataRange.LastCell().GetText());

            table.SetShowHeaderRow();
            IXLRangeRow headerRow = table.HeadersRow();
            Assert.AreNotEqual(null, headerRow);
            Assert.AreEqual("Categories", headerRow.Cell(1).GetText());

            table.SetShowHeaderRow(false);

            ws.FirstCell().SetValue("x");

            table.SetShowHeaderRow();

            Assert.AreEqual("x", ws.FirstCell().GetText());
            Assert.AreEqual("Categories", ws.Cell("A2").GetText());
            Assert.AreNotEqual(null, headerRow);
            Assert.AreEqual("A", table.DataRange.FirstRow().Field("Categories").GetText());
            Assert.AreEqual("C", table.DataRange.LastRow().Field("Categories").GetText());
            Assert.AreEqual("A", table.DataRange.FirstCell().GetText());
            Assert.AreEqual("C", table.DataRange.LastCell().GetText());
        }
    }

    [TestCase("Amount")]
    [TestCase("AMOUNT")]
    [TestCase("amount")]
    public void FieldNamesOfXLTableAreCaseInsensitive(string fieldName)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLTable table = ws.Cell("A1").InsertTable(new[] { new { Amount = 1 } });

        IXLTableField expectedField = table.Field(0);
        Assert.AreSame(expectedField, table.Field(fieldName));
    }

    [Test]
    public void ChangeFieldName()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").SetValue("FName").CellBelow().SetValue("John");

            ws.Cell("B1").SetValue("LName").CellBelow().SetValue("Doe");

            IXLTable tbl = ws.RangeUsed().CreateTable();
            string nameBefore = tbl.Field(tbl.Fields.Last().Index).Name;
            tbl.Field(tbl.Fields.Last().Index).Name = "LastName";
            string nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;

            string cellValue = ws.Cell("B1").GetText();

            Assert.AreEqual("LName", nameBefore);
            Assert.AreEqual("LastName", nameAfter);
            Assert.AreEqual("LastName", cellValue);

            tbl.ShowHeaderRow = false;
            tbl.Field(tbl.Fields.Last().Index).Name = "LastNameChanged";
            nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;
            Assert.AreEqual("LastNameChanged", nameAfter);

            tbl.SetShowHeaderRow(true);
            nameAfter = (String)tbl.Cell("B1").Value;
            Assert.AreEqual("LastNameChanged", nameAfter);

            IXLTableField field = tbl.Field("LastNameChanged");
            Assert.AreEqual("LastNameChanged", field.Name);

            tbl.Cell(1, 1).Value = "FirstName";
            Assert.AreEqual("FirstName", tbl.Field(0).Name);
        }
    }

    [Test]
    public void CanDeleteTableColumn()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLTable table = ws.FirstCell().InsertTable(l);

            table.Column("C").Delete();

            Assert.AreEqual(3, table.Fields.Count());

            Assert.AreEqual("FirstColumn", table.Fields.First().Name);
            Assert.AreEqual(0, table.Fields.First().Index);

            Assert.AreEqual("UnOrderedColumn", table.Fields.Last().Name);
            Assert.AreEqual(2, table.Fields.Last().Index);
        }
    }

    [Test]
    public void TestFieldCellTypes()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLTable table = ws.Cell("B2").InsertTable(l);

            Assert.AreEqual(4, table.Fields.Count());

            Assert.AreEqual("B2", table.Field(0).HeaderCell.Address.ToString());
            Assert.AreEqual("C2", table.Field(1).HeaderCell.Address.ToString());
            Assert.AreEqual("D2", table.Field(2).HeaderCell.Address.ToString());
            Assert.AreEqual("E2", table.Field(3).HeaderCell.Address.ToString());

            Assert.IsNull(table.Field(0).TotalsCell);
            Assert.IsNull(table.Field(1).TotalsCell);
            Assert.IsNull(table.Field(2).TotalsCell);
            Assert.IsNull(table.Field(3).TotalsCell);

            table.SetShowTotalsRow();

            Assert.AreEqual("B5", table.Field(0).TotalsCell.Address.ToString());
            Assert.AreEqual("C5", table.Field(1).TotalsCell.Address.ToString());
            Assert.AreEqual("D5", table.Field(2).TotalsCell.Address.ToString());
            Assert.AreEqual("E5", table.Field(3).TotalsCell.Address.ToString());

            IXLTableField field = table.Fields.Last();

            Assert.AreEqual("E2:E5", field.Column.RangeAddress.ToString());
            Assert.AreEqual("E3", field.DataCells.First().Address.ToString());
            Assert.AreEqual("E4", field.DataCells.Last().Address.ToString());
        }
    }

    [Test]
    public void CanDeleteTable()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().InsertTable(l);
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLTable table = ws.Tables.First();

                ws.Tables.Remove(table.Name);
                Assert.AreEqual(0, ws.Tables.Count());
                wb.Save();
            }
        }
    }

    [Test]
    public void TableNameCannotBeValidCellName()
    {
        DataTable dt = new("sheet1");
        dt.Columns.Add("Patient", typeof(string));
        dt.Rows.Add("David");

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<InvalidOperationException>(() =>
                ws.Cell(1, 1).InsertTable(dt, "May2019")
            );
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "A1"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "R1C2"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "r3c2"));
            Assert.Throws<InvalidOperationException>(() =>
                ws.Cell(1, 1).InsertTable(dt, "R2C33333")
            );
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "RC"));
        }
    }

    [Test]
    public void TableNameSetWhenAddingWorksheetWithDataTable()
    {
        DataTable dt = new("sheet1");
        dt.Columns.Add("Patient", typeof(string));
        dt.Rows.Add("David");

        using (XLWorkbook wb = new())
        {
            // Generated table name is used and should not be an issue
            Assert.DoesNotThrow(() => wb.AddWorksheet(dt, "t1"));
        }

        using (XLWorkbook wb = new())
        {
            // Should pass because t1 is a valid sheet name, and is not used for the tableName
            Assert.DoesNotThrow(() => wb.AddWorksheet(dt, "t1", "table1"));

            Assert.AreEqual(1, wb.Worksheets.Count);
            Assert.AreEqual(1, wb.Worksheet(1).Tables.Count());
        }
    }

    [Test]
    public void CanDeleteTableField()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLTable table = ws.Cell("B2").InsertTable(l);

            Assert.AreEqual("B2:E4", table.RangeAddress.ToString());

            table.Field("SomeFieldNotProperty").Delete();

            Assert.AreEqual(3, table.Fields.Count());

            Assert.AreEqual("FirstColumn", table.Fields.First().Name);
            Assert.AreEqual(0, table.Fields.First().Index);

            Assert.AreEqual("UnOrderedColumn", table.Fields.Last().Name);
            Assert.AreEqual(2, table.Fields.Last().Index);

            Assert.AreEqual("B2:D4", table.RangeAddress.ToString());
        }
    }

    [Test]
    public void CanDeleteTableRows()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
            new()
            {
                Column1 = "e",
                Column2 = "f",
                MyField = 6,
                UnOrderedColumn = 555,
            },
            new()
            {
                Column1 = "g",
                Column2 = "h",
                MyField = 7,
                UnOrderedColumn = 333,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLTable table = ws.Cell("B2").InsertTable(l);

            Assert.AreEqual("B2:E6", table.RangeAddress.ToString());

            table.DataRange.Rows(3, 4).Delete();

            Assert.AreEqual(2, table.DataRange.Rows().Count());

            Assert.AreEqual("b", table.DataRange.FirstCell().Value);
            Assert.AreEqual(777, table.DataRange.LastCell().Value);

            Assert.AreEqual("B2:E4", table.RangeAddress.ToString());
        }
    }

    [Test]
    public void OverlappingTablesThrowsException()
    {
        DataTable dt = new("sheet1");
        dt.Columns.Add("col1", typeof(string));
        dt.Columns.Add("col2", typeof(double));

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(dt, true);
            Assert.Throws<InvalidOperationException>(() =>
                ws.FirstCell().CellRight().InsertTable(dt, true)
            );
        }
    }

    [Test]
    public void OverwritingTableHeaders()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLTable table = ws.Cell("A1")
            .InsertTable(new object[] { ("Header 1", "Header 2"), (1, 2) }, true);

        // Overwrite the headers of the table with non-string values
        ws.Cell("A1").InsertData(new object[] { (XLError.IncompatibleValue, 7) });

        // The non-string data inserted to headers were converted to strings and used as a field names.
        Assert.AreEqual("#VALUE!", table.Field(0).Name);
        Assert.AreEqual("#VALUE!", ws.Cell("A1").Value);
        Assert.AreEqual("7", table.Field(1).Name);
        Assert.AreEqual("7", ws.Cell("B1").Value);
    }

    [Test]
    public void OverwritingTableTotalsRow()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var data1 = Enumerable
                .Range(1, 10)
                .Select(i => new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new String('a', i),
                });

            IXLTable table = ws.FirstCell()
                .InsertTable(data1, true)
                .SetShowHeaderRow()
                .SetShowTotalsRow();
            table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

            var data2 = Enumerable
                .Range(1, 20)
                .Select(i => new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new String('b', i),
                    Int = 64 + i,
                });

            ws.FirstCell().CellBelow().InsertData(data2);

            table.Fields.ForEach(f =>
                Assert.AreEqual(XLTotalsRowFunction.None, f.TotalsRowFunction)
            );

            Assert.AreEqual("11", table.Field(0).TotalsRowLabel);
            Assert.AreEqual("K", table.Field(1).TotalsRowLabel);
            Assert.AreEqual("bbbbbbbbbbb", table.Field(2).TotalsRowLabel);
        }
    }

    [Test]
    public void TableRenameTests()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLTable table1 = ws.FirstCell().InsertTable(l);
            IXLTable table2 = ws.Cell("A10").InsertTable(l);

            Assert.AreEqual("Table1", table1.Name);
            Assert.AreEqual("Table2", table2.Name);

            table1.Name = "table1";
            Assert.AreEqual("table1", table1.Name);

            table1.Name = "_table1";
            Assert.AreEqual("_table1", table1.Name);

            table1.Name = "\\table1";
            Assert.AreEqual("\\table1", table1.Name);

            Assert.Throws<ArgumentException>(() => table1.Name = "");
            Assert.Throws<ArgumentException>(() => table1.Name = "R");
            Assert.Throws<ArgumentException>(() => table1.Name = "C");
            Assert.Throws<ArgumentException>(() => table1.Name = "r");
            Assert.Throws<ArgumentException>(() => table1.Name = "c");

            Assert.Throws<ArgumentException>(() => table1.Name = "123");
            Assert.Throws<ArgumentException>(() => table1.Name = new String('A', 256));

            Assert.Throws<ArgumentException>(() => table1.Name = "Table2");
            Assert.Throws<ArgumentException>(() => table1.Name = "TABLE2");
        }
    }

    [Test]
    public void CanResizeTable()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var data1 = Enumerable
                .Range(1, 10)
                .Select(i => new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new String('a', i),
                });

            IXLTable table = ws.FirstCell()
                .InsertTable(data1, true)
                .SetShowHeaderRow()
                .SetShowTotalsRow();
            table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

            var data2 = Enumerable
                .Range(1, 10)
                .Select(i => new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new String('b', i),
                    Integer = 64 + i,
                });

            ws.FirstCell().CellBelow().InsertData(data2);
            table.Resize(table.FirstCell().Address, table.AsRange().LastCell().CellRight().Address);

            Assert.AreEqual(4, table.Fields.Count());

            Assert.AreEqual("Column4", table.Field(3).Name);

            ws.Cell("D1").Value = "Integer";
            Assert.AreEqual("Integer", table.Field(3).Name);
        }
    }

    [Test]
    public void TableAsDynamicEnumerable()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLTable table = ws.FirstCell().InsertTable(l);

            foreach (dynamic d in table.AsDynamicEnumerable())
            {
                Assert.DoesNotThrow(() =>
                {
                    object value;
                    value = d.FirstColumn;
                    value = d.SecondColumn;
                    value = d.UnOrderedColumn;
                    value = d.SomeFieldNotProperty;
                });
            }
        }
    }

    [Test]
    public void TableAsDotNetDataTable()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            DataTable table = ws.FirstCell().InsertTable(l).AsNativeDataTable();

            Assert.AreEqual(4, table.Columns.Count);
            Assert.AreEqual("FirstColumn", table.Columns[0].ColumnName);
            Assert.AreEqual("SecondColumn", table.Columns[1].ColumnName);
            Assert.AreEqual("SomeFieldNotProperty", table.Columns[2].ColumnName);
            Assert.AreEqual("UnOrderedColumn", table.Columns[3].ColumnName);

            Assert.AreEqual(typeof(String), table.Columns[0].DataType);
            Assert.AreEqual(typeof(String), table.Columns[1].DataType);
            Assert.AreEqual(typeof(Double), table.Columns[2].DataType);
            Assert.AreEqual(typeof(Double), table.Columns[3].DataType);

            DataRow dr = table.Rows[0];
            Assert.AreEqual("b", dr["FirstColumn"]);
            Assert.AreEqual("a", dr["SecondColumn"]);
            Assert.AreEqual(4, dr["SomeFieldNotProperty"]);
            Assert.AreEqual(999, dr["UnOrderedColumn"]);

            dr = table.Rows[1];
            Assert.AreEqual("d", dr["FirstColumn"]);
            Assert.AreEqual("c", dr["SecondColumn"]);
            Assert.AreEqual(5, dr["SomeFieldNotProperty"]);
            Assert.AreEqual(777, dr["UnOrderedColumn"]);
        }
    }

    [Test]
    public void TestTableCellTypes()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var data1 = Enumerable
                .Range(1, 10)
                .Select(i => new
                {
                    Index = i,
                    Character = Convert.ToChar(64 + i),
                    String = new String('a', i),
                });

            IXLTable table = ws.FirstCell()
                .InsertTable(data1, true)
                .SetShowHeaderRow()
                .SetShowTotalsRow();
            table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

            Assert.AreEqual(XLTableCellType.Header, table.HeadersRow().Cell(1).TableCellType());
            Assert.AreEqual(
                XLTableCellType.Data,
                table.HeadersRow().Cell(1).CellBelow().TableCellType()
            );
            Assert.AreEqual(XLTableCellType.Total, table.TotalsRow().Cell(1).TableCellType());
            Assert.AreEqual(XLTableCellType.None, ws.Cell("Z100").TableCellType());
        }
    }

    [Test]
    public void TotalsFunctionsOfHeadersWithWeirdCharacters()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l, false);

            // Give the headings weird names (i.e. spaces, hashes, single quotes
            ws.Cell("A1").Value = "ABCD    ";
            ws.Cell("B1").Value = "   #BCD";
            ws.Cell("C1").Value = "   as'df   ";
            ws.Cell("D1").Value = "Normal";

            IXLTable table = ws.RangeUsed().CreateTable();
            Assert.IsNotNull(table);

            table.ShowTotalsRow = true;
            table.Field(0).TotalsRowFunction = XLTotalsRowFunction.Count;
            table.Field(1).TotalsRowFunction = XLTotalsRowFunction.Count;
            table.Field(2).TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field(3).TotalsRowFunction = XLTotalsRowFunction.Sum;

            Assert.AreEqual("SUBTOTAL(103,Table1[[ABCD    ]])", table.Field(0).TotalsRowFormulaA1);
            Assert.AreEqual("SUBTOTAL(103,Table1[[   '#BCD]])", table.Field(1).TotalsRowFormulaA1);
            Assert.AreEqual(
                "SUBTOTAL(109,Table1[[   as''df   ]])",
                table.Field(2).TotalsRowFormulaA1
            );
            Assert.AreEqual("SUBTOTAL(109,[Normal])", table.Field(3).TotalsRowFormulaA1);
        }
    }

    [Test]
    public void CannotCreateDuplicateTablesOverSameRange()
    {
        List<TestObjectWithAttributes> l =
        [
            new()
            {
                Column1 = "a",
                Column2 = "b",
                MyField = 4,
                UnOrderedColumn = 999,
            },
            new()
            {
                Column1 = "c",
                Column2 = "d",
                MyField = 5,
                UnOrderedColumn = 777,
            },
        ];

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.Throws<InvalidOperationException>(() => ws.RangeUsed().CreateTable());
        }
    }

    [Test]
    public void CannotCreateTableOverExistingAutoFilter()
    {
        using XLWorkbook wb = new();

        var data = Enumerable.Range(1, 10).Select(i => new { Index = i, String = $"String {i}" });

        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell().InsertTable(data, createTable: false);
        ws.RangeUsed().SetAutoFilter().Column(1).AddFilter(5);

        Assert.Throws<InvalidOperationException>(() => ws.RangeUsed().CreateTable());
    }

    [Test]
    public void CopyTableSameWorksheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");

        IXLTable table = ws1.Range("A1:C2").AsTable();

        TestDelegate action = () => table.CopyTo(ws1);

        Assert.Throws(typeof(InvalidOperationException), action);
    }

    [Test]
    public void CanInsertDateTimeOffset()
    {
        DateTimeOffset now = DateTimeOffset.Now;

        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        ws1.FirstCell().InsertTable(new[] { new { TimeStamp = now } });

        // C# Supports 7 digits milliseconds, but excel only 3
        const string format = "yyyy-MM-dd HH:mm:ss.fff";

        string actual = ws1.Cell("A2").GetDateTime().ToString(format);
        string expected = now.DateTime.ToString(format);
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void CopyDetachedTableDifferentWorksheets()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Custom column 1";
        ws1.Cell("B1").Value = "Custom column 2";
        ws1.Cell("C1").Value = "Custom column 3";
        ws1.Cell("A2").Value = "Value 1";
        ws1.Cell("B2").Value = 123.45;
        ws1.Cell("C2").Value = new DateTime(2018, 5, 10);
        IXLTable original = ws1.Range("A1:C2").AsTable("Detached table");
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");

        IXLTable copy = original.CopyTo(ws2);

        Assert.AreEqual(0, ws1.Tables.Count()); // We did not add it
        Assert.AreEqual(1, ws2.Tables.Count());

        AssertTablesAreEqual(original, copy);

        Assert.AreEqual("Sheet2!A1:C2", copy.RangeAddress.ToString(XLReferenceStyle.A1, true));
        Assert.AreEqual("Custom column 1", ws2.Cell("A1").Value);
        Assert.AreEqual("Custom column 2", ws2.Cell("B1").Value);
        Assert.AreEqual("Custom column 3", ws2.Cell("C1").Value);
        Assert.AreEqual("Value 1", ws2.Cell("A2").Value);
        Assert.AreEqual(123.45, (double)ws2.Cell("B2").Value, XLHelper.Epsilon);
        Assert.AreEqual(new DateTime(2018, 5, 10), ws2.Cell("C2").Value);
    }

    [Test]
    public void CopyTableDifferentWorksheets()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Custom column 1";
        ws1.Cell("B1").Value = "Custom column 2";
        ws1.Cell("C1").Value = "Custom column 3";
        ws1.Cell("A2").Value = "Value 1";
        ws1.Cell("B2").Value = 123.45;
        ws1.Cell("C2").Value = new DateTime(2018, 5, 10);
        IXLTable original = ws1.Range("A1:C2").AsTable("Attached table");
        ws1.Tables.Add(original);
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");

        original.CopyTo(ws2);

        Assert.AreEqual(1, ws1.Tables.Count());
        Assert.AreEqual(1, ws2.Tables.Count());

        IXLTable copy = ws2.Tables.First();

        AssertTablesAreEqual(original, copy);

        Assert.AreEqual("Sheet2!A1:C2", copy.RangeAddress.ToString(XLReferenceStyle.A1, true));
        Assert.AreEqual("Custom column 1", ws2.Cell("A1").Value);
        Assert.AreEqual("Custom column 2", ws2.Cell("B1").Value);
        Assert.AreEqual("Custom column 3", ws2.Cell("C1").Value);
        Assert.AreEqual("Value 1", ws2.Cell("A2").Value);
        Assert.AreEqual(123.45, (double)ws2.Cell("B2").Value, XLHelper.Epsilon);
        Assert.AreEqual(new DateTime(2018, 5, 10), ws2.Cell("C2").Value);
    }

    [Test]
    public void NewTableHasNullRelId()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Custom column 1";
                ws.Cell("B1").Value = "Custom column 2";
                ws.Cell("C1").Value = "Custom column 3";
                ws.Cell("A2").Value = "Value 1";
                ws.Cell("B2").Value = 123.45;
                ws.Cell("C2").Value = new DateTime(2018, 5, 10);
                IXLTable original = ws.Range("A1:C2").CreateTable("Attached table");

                Assert.AreEqual(1, ws.Tables.Count());
                Assert.IsNull((original as XLTable).RelId);

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet2");
                IXLTable original = wb.Worksheets.First().Tables.First();

                Assert.IsNotNull((original as XLTable).RelId);

                IXLTable copy = original.CopyTo(ws);

                Assert.AreEqual(1, ws.Tables.Count());
                Assert.IsNull((copy as XLTable).RelId);

                AssertTablesAreEqual(original, copy);

                Assert.AreEqual(
                    "Sheet2!A1:C2",
                    copy.RangeAddress.ToString(XLReferenceStyle.A1, true)
                );
                Assert.AreEqual("Custom column 1", ws.Cell("A1").Value);
                Assert.AreEqual("Custom column 2", ws.Cell("B1").Value);
                Assert.AreEqual("Custom column 3", ws.Cell("C1").Value);
                Assert.AreEqual("Value 1", ws.Cell("A2").Value);
                Assert.AreEqual(123.45, (double)ws.Cell("B2").Value, XLHelper.Epsilon);
                Assert.AreEqual(new DateTime(2018, 5, 10), ws.Cell("C2").Value);
            }
        }
    }

    [Test]
    public void CopyTableWithoutData()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Custom column 1";
        ws1.Cell("B1").Value = "Custom column 2";
        ws1.Cell("C1").Value = "Custom column 3";
        ws1.Cell("A2").Value = "Value 1";
        ws1.Cell("B2").Value = 123.45;
        ws1.Cell("C2").Value = new DateTime(2018, 5, 10);
        IXLTable original = ws1.Range("A1:C2").AsTable("Attached table");
        ws1.Tables.Add(original);
        XLWorksheet? ws2 = wb.Worksheets.Add("Sheet2") as XLWorksheet;

        IXLTable copy = (original as XLTable).CopyTo(ws2, false);

        AssertTablesAreEqual(original, copy);

        Assert.AreEqual("Sheet2!A1:C2", copy.RangeAddress.ToString(XLReferenceStyle.A1, true));
        Assert.AreEqual("Custom column 1", ws2.Cell("A1").Value);
        Assert.AreEqual("Custom column 2", ws2.Cell("B1").Value);
        Assert.AreEqual("Custom column 3", ws2.Cell("C1").Value);
        Assert.AreEqual(Blank.Value, ws2.Cell("A2").Value);
        Assert.AreEqual(Blank.Value, ws2.Cell("B2").Value);
        Assert.AreEqual(Blank.Value, ws2.Cell("C2").Value);
    }

    [Test]
    public void SavingTableWithNullDataRangeThrowsException()
    {
        using (MemoryStream ms = new())
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var data = Enumerable
                .Range(1, 10)
                .Select(i => new
                {
                    Number = i,
                    NumberString = String.Concat("Number", i.ToString()),
                });

            IXLTable table = ws.FirstCell().InsertTable(data).SetShowTotalsRow();

            table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Count;

            table
                .DataRange.Rows()
                .OrderByDescending(r => r.RowNumber())
                .ToList()
                .ForEach(r => r.WorksheetRow().Delete());

            Assert.IsNull(table.DataRange);
            Assert.Throws<EmptyTableException>(() => wb.SaveAs(ms));
        }
    }

    [Test]
    public void SaveTotalsRowLabelCellWithSstIdMatchingTheLabel()
    {
        // Issue #2602 test. The totals row wasn't saved with compact SST ID from file, but with a memory SST that has holes.
        TestHelper.CreateAndCompare(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();
                ws.Cell("A1").Value = "Dummy1"; // First inserted text - index=0, reference count = 1
                ws.Cell("A2").Value = "Dummy2"; // Second inserted text - index=1, reference count = 1
                ws.Cell("A3").Value = "Dummy3"; // Third inserted text - index=2, reference count = 1
                ws.Cell("A4").Value = "Text"; // Fourth inserted text - index=3, reference count = 1
                IXLTable table = ws.Cell("A5").InsertTable([("Text", 17)]); // Also inserts header Item1 and Item2.
                table.ShowTotalsRow = true;
                table.Field(0).TotalsRowLabel = "Text"; // reference count = 3

                // Remove "Dummy*" text. That way, the "Text", "Item1" and "Item2" will be in index 0..2 that were occupied by Dummy*
                // Ensure that cell in total row label A7 references "Text" SST ID
                ws.Range("A1:A3").Value = Blank.Value;
            },
            @"Other\Tables\TotalRowSstId.xlsx"
        );
    }

    [Test]
    public void CanCreateTableWithWhiteSpaceColumnHeaders()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue("Header");
            ws.Cell("B1").SetValue(new string(' ', 1));
            ws.Cell("C1").SetValue(new string(' ', 2));
            ws.Cell("D1").SetValue(new string(' ', 3));

            IXLTable table = ws.Range("A1:E3").CreateTable("Table1");

            Assert.AreEqual("Header", table.Field(0).Name);
            Assert.AreEqual(new string(' ', 1), table.Field(1).Name);
            Assert.AreEqual(new string(' ', 2), table.Field(2).Name);
            Assert.AreEqual(new string(' ', 3), table.Field(3).Name);
            Assert.AreEqual("Column5", table.Field(4).Name);
        }
    }

    [Test]
    public void TableNotFound()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Table("dummy"));
            Assert.Throws<ArgumentOutOfRangeException>(() => wb.Table("dummy"));
        }
    }

    [Test]
    public void SecondTableOnNewSheetHasUniqueName()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet();
            IXLTable t1 = ws1.FirstCell()
                .InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
            Assert.AreEqual("Table1", t1.Name);

            IXLWorksheet ws2 = wb.AddWorksheet();
            IXLTable t2 = ws2.FirstCell()
                .InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
            Assert.AreEqual("Table2", t2.Name);
        }
    }

    private static void AssertTablesAreEqual(IXLTable table1, IXLTable table2)
    {
        Assert.AreEqual(
            table1.RangeAddress.ToString(XLReferenceStyle.A1, false),
            table2.RangeAddress.ToString(XLReferenceStyle.A1, false)
        );
        Assert.AreEqual(table1.Fields.Count(), table2.Fields.Count());
        for (int j = 0; j < table1.Fields.Count(); j++)
        {
            IXLTableField originalField = table1.Fields.ElementAt(j);
            IXLTableField copyField = table2.Fields.ElementAt(j);
            Assert.AreEqual(originalField.Name, copyField.Name);
            if (table1.ShowTotalsRow)
            {
                Assert.AreEqual(originalField.TotalsRowFormulaA1, copyField.TotalsRowFormulaA1);
                Assert.AreEqual(originalField.TotalsRowFunction, copyField.TotalsRowFunction);
            }
        }

        Assert.AreEqual(table1.Name, table2.Name);
        Assert.AreEqual(table1.ShowAutoFilter, table2.ShowAutoFilter);
        Assert.AreEqual(table1.ShowColumnStripes, table2.ShowColumnStripes);
        Assert.AreEqual(table1.ShowHeaderRow, table2.ShowHeaderRow);
        Assert.AreEqual(table1.ShowRowStripes, table2.ShowRowStripes);
        Assert.AreEqual(table1.ShowTotalsRow, table2.ShowTotalsRow);
        Assert.AreEqual(table1.Theme, table2.Theme);
    }

    //TODO: Delete table (not underlying range)
}
