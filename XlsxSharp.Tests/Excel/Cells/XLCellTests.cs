using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.Cells;

public class XlCellTests
{
    [SuppressMessage("ReSharper", "RedundantCast")]
    private static readonly object[] AllNumberTypes =
    [
        (sbyte)1,
        (byte)2,
        (short)3,
        (ushort)4,
        (int)5,
        (uint)6,
        (long)7,
        (ulong)8,
        (float)9.5f,
        (double)10.75,
        (decimal)11.875m,
    ];

    [Test]
    public void CellsUsed()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Cell(1, 1);
        ws.Cell(2, 2);
        int count = ws.Range("A1:B2").CellsUsed().Count();
        ClassicAssert.AreEqual(0, count);
    }

    [Test]
    public void CellsUsedIncludeFormatDifferentFromInheritedFormat1()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Row(3).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Column(3).Style.Fill.BackgroundColor = XLColor.Blue;
        ws.Cell(2, 2).Value = "ASDF";
        string? range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
        ClassicAssert.AreEqual("B2:C3", range);
    }

    [Test]
    public void CellsUsedIncludeFormatDifferentFromInheritedFormat2()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Row(2).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Column(2).Style.Fill.BackgroundColor = XLColor.Blue;
        ws.Cell(3, 3).Value = "ASDF";
        string? range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
        ClassicAssert.AreEqual("B2:C3", range);
    }

    [Test]
    public void CellsUsedIncludeStyles3()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRange? range = ws.RangeUsed(XLCellsUsedOptions.All);
        ClassicAssert.AreEqual(null, range);
    }

    [Test]
    public void CellUsedIncludesSparklines()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:E4").Value = 1;
        ws.SparklineGroups.Add("B2", "C3:E3");
        ws.SparklineGroups.Add("F5", "C4:E4");

        string? range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
        ClassicAssert.AreEqual("B2:F5", range);
    }

    [Test]
    public void GetValueNullable()
    {
        IXLCell cell = new XLWorkbook().AddWorksheet().FirstCell();

        ClassicAssert.IsNull(cell.Clear().GetValue<double?>());
        ClassicAssert.AreEqual(1.5, cell.SetValue(1.5).GetValue<double?>());
        ClassicAssert.AreEqual(2, cell.SetValue(2).GetValue<int?>());
        ClassicAssert.IsNull(cell.SetValue(Blank.Value).GetValue<double?>());
        ClassicAssert.Throws<InvalidCastException>(() => cell.SetValue("text").GetValue<double?>());
    }

    [Test]
    public void InsertData1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" });
        ClassicAssert.AreEqual("Sheet1!B2:B4", range.ToString());
    }

    [Test]
    public void InsertDataDoesntTransposeDataOnFalseFlag()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, false);
        ClassicAssert.AreEqual("Sheet1!B2:B4", range.ToString());
    }

    [Test]
    public void InsertDataTransposesDataOnTrueFlag()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, true);
        ClassicAssert.AreEqual("Sheet1!B2:D2", range.ToString());
    }

    [Test]
    public void InsertDataDifferentTypes()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        object[] values = ["Text", 45, DateTime.Today, true, "More text"];

        ws.FirstCell().InsertData(values);

        ClassicAssert.AreEqual("Text", ws.FirstCell().GetString());
        ClassicAssert.AreEqual(45, ws.Cell("A2").GetDouble());
        ClassicAssert.AreEqual(DateTime.Today, ws.Cell("A3").GetDateTime());
        ClassicAssert.AreEqual(true, ws.Cell("A4").GetBoolean());
        ClassicAssert.AreEqual("More text", ws.Cell("A5").GetString());
        ClassicAssert.IsTrue(ws.Cell("A6").IsEmpty());
    }

    [Test]
    public void InsertDataWithGuids()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.FirstCell()
            .InsertData(Enumerable.Range(1, 20).Select(i => new { Guid = Guid.NewGuid() }));

        ClassicAssert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
        ClassicAssert.AreEqual(Guid.NewGuid().ToString().Length, ws.FirstCell().GetText().Length);
    }

    [Test]
    public void InsertDataWithNulls()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");

        DataTable table = new();
        table.TableName = "Patients";
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
        table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
        table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
        table.Rows.Add(21, "Combivent", DBNull.Value, new DateTime(2000, 1, 4));
        table.Rows.Add(100, "Dilantin", "Melanie", DBNull.Value);

        ws.FirstCell().InsertData(table);

        ClassicAssert.AreEqual(25, ws.Cell("A1").Value);
        ClassicAssert.AreEqual("", ws.Cell("C4").Value);
        ClassicAssert.AreEqual("", ws.Cell("D5").Value);
    }

    [Test]
    public void InsertDataWithNullsIEnumerable()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");

        List<DateTime?> dateTimeList =
        [
            new DateTime(2000, 1, 1),
            new DateTime(2000, 1, 2),
            new DateTime(2000, 1, 3),
            new DateTime(2000, 1, 4),
            null,
        ];

        ws.FirstCell().InsertData(dateTimeList);

        ClassicAssert.AreEqual(new DateTime(2000, 1, 1), ws.Cell("A1").GetDateTime());
        ClassicAssert.AreEqual(Blank.Value, ws.Cell("A5").Value);
    }

    [Test]
    public void InsertDataAllNumberTypesAreInsertedAsNumbers()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add();

        ws.FirstCell().InsertData(AllNumberTypes);

        for (int row = 1; row <= AllNumberTypes.Length; ++row)
        {
            object expectedValue = Convert.ChangeType(AllNumberTypes[row - 1], typeof(double));
            XLCellValue actualValue = ws.Cell(row, 1).Value;
            ClassicAssert.AreEqual(expectedValue, actualValue);
        }
    }

    [Test]
    public void InsertTableAllNumberTypesAreInsertedAsNumbers()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add();

        DataTable table = new("Numbers");
        foreach (object number in AllNumberTypes)
        {
            Type numberType = number.GetType();
            table.Columns.Add(numberType.Name, numberType);
        }

        table.Rows.Add(AllNumberTypes);

        ws.FirstCell().InsertTable(table);

        for (int column = 1; column <= AllNumberTypes.Length; ++column)
        {
            object expectedValue = Convert.ChangeType(AllNumberTypes[column - 1], typeof(double));
            XLCellValue actualValue = ws.Cell(2, column).Value;
            ClassicAssert.AreEqual(expectedValue, actualValue);
        }
    }

    [Test]
    public void IsEmpty1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        bool actual = cell.IsEmpty();
        bool expected = true;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void IsEmpty2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        bool actual = cell.IsEmpty(XLCellsUsedOptions.All);
        bool expected = true;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void IsEmpty3()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        bool actual = cell.IsEmpty();
        bool expected = true;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void IsEmpty4()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        bool actual = cell.IsEmpty(XLCellsUsedOptions.AllContents);
        bool expected = true;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void IsEmpty5()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        bool actual = cell.IsEmpty(XLCellsUsedOptions.All);
        bool expected = false;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void IsEmpty6()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.Value = "X";
        bool actual = cell.IsEmpty();
        bool expected = false;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void NaNIsNotANumber()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1");
        cell.Value = "NaN";

        ClassicAssert.AreNotEqual(XLDataType.Number, cell.DataType);
    }

    [Test]
    public void NanIsNotANumber()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1");
        cell.Value = "Nan";

        ClassicAssert.AreNotEqual(XLDataType.Number, cell.DataType);
    }

    [Test]
    public void TryGetValueBooleanBad()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue("ABC");
        bool success = cell.TryGetValue(out bool _);
        ClassicAssert.IsFalse(success);
    }

    [Test]
    public void TryGetValueBooleanFalse()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue(false);
        bool success = cell.TryGetValue(out bool outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.IsFalse(outValue);
    }

    [Test]
    public void TryGetValueBooleanFalseText()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue("False");
        bool success = cell.TryGetValue(out bool outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.IsFalse(outValue);
    }

    [Test]
    public void TryGetValueBooleanTrue()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue(true);
        bool success = cell.TryGetValue(out bool outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.IsTrue(outValue);
    }

    [Test]
    public void TryGetValueBooleanTrueText()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue("True");
        bool success = cell.TryGetValue(out bool outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.IsTrue(outValue);
    }

    [Test]
    public void TryGetValueDateTimeGood2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        bool success = ws.Cell("A1")
            .SetFormulaA1("=TODAY() + 10")
            .TryGetValue(out DateTime outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(DateTime.Today.AddDays(10), outValue);
    }

    [Test]
    public void TryGetValueDateTimeBadButFormulaGood()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        bool success = ws.Cell("A1")
            .SetFormulaA1("=\"44\"&\"020\"")
            .TryGetValue(out DateTime outValue);
        ClassicAssert.IsFalse(success);

        ws.Cell("B1").SetFormulaA1("=A1+1");

        success = ws.Cell("B1").TryGetValue(out outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(new DateTime(2020, 07, 09), outValue);
    }

    [Test]
    public void TryGetValueDateTimeBadString()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        string date = "ABC";
        bool success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime _);
        ClassicAssert.IsFalse(success);
    }

    [Test]
    public void TryGetValueDateTimeSerialDateTimeOutsideRange()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        int serialDateTimeOutsideRange = 5545454;
        ws.FirstCell().SetValue(serialDateTimeOutsideRange);
        bool success = ws.FirstCell().TryGetValue(out DateTime _);
        ClassicAssert.IsFalse(success);
    }

    [Test]
    public void TryGetValueEnumGood()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet();
        ClassicAssert.IsTrue(
            ws.FirstCell()
                .SetValue(nameof(NumberStyles.AllowCurrencySymbol))
                .TryGetValue(out NumberStyles value)
        );
        ClassicAssert.AreEqual(NumberStyles.AllowCurrencySymbol, value);

        // Nullable alternative
        ClassicAssert.IsTrue(
            ws.FirstCell()
                .SetValue(nameof(NumberStyles.AllowCurrencySymbol))
                .TryGetValue(out NumberStyles? value2)
        );
        ClassicAssert.AreEqual(NumberStyles.AllowCurrencySymbol, value2);
    }

    [Test]
    public void TryGetValueEnumBadString()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet();
        ClassicAssert.IsFalse(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles _));
        ClassicAssert.IsFalse(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles? _));
    }

    [Test]
    public void TryGetValueTimeSpanBadString()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        string timeSpan = "ABC";
        bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan _);
        ClassicAssert.IsFalse(success);
    }

    [Test]
    public void TryGetValueTimeSpanGood()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        TimeSpan timeSpan = new(1, 1, 1);
        bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(timeSpan, outValue);
    }

    [Test]
    public void TryGetValueTimeSpanGood2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        bool success = ws.Cell("A1")
            .SetValue(0.0034722222222222199)
            .TryGetValue(out TimeSpan outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(TimeSpan.FromMinutes(5), outValue);
    }

    [Test]
    public void TryGetValueTimeSpanGoodLarge()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        TimeSpan timeSpan = TimeSpan.FromMilliseconds((double)int.MaxValue + 1);
        bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(timeSpan, outValue);
    }

    [Test]
    public void TryGetValueTimeSpanGoodFromText()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        bool success = ws.Cell("A1").SetValue("300:14:50.453").TryGetValue(out TimeSpan outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(new TimeSpan(12, 12, 14, 50, 453), outValue);
    }

    [Test]
    public void TryGetValueSbyteBad2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue("255");
        bool success = cell.TryGetValue(out sbyte _);
        ClassicAssert.IsFalse(success);
    }

    [Test]
    public void TryGetValueSbyteGood()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell("A1").SetValue(5);
        bool success = cell.TryGetValue(out sbyte outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual(5, outValue);
    }

    [Test]
    public void TryGetValueUnicodeString()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");

        bool success;
        string outValue;

        success = ws.Cell("A1").SetValue("Site_x0020_Column_x0020_Test").TryGetValue(out outValue);
        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual("Site Column Test", outValue);

        success = ws.Cell("A1")
            .SetValue("Site_x005F_x0020_Column_x005F_x0020_Test")
            .TryGetValue(out outValue);

        ClassicAssert.IsTrue(success);
        ClassicAssert.AreEqual("Site_x005F_x0020_Column_x005F_x0020_Test", outValue);
    }

    [Test]
    public void TryGetValueNullable()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Clear();
        ws.Cell("A2").SetValue(1.5);
        ws.Cell("A3").SetValue(2.5.ToString(CultureInfo.CurrentCulture));
        ws.Cell("A4").SetValue("text");

        ClassicAssert.IsTrue(ws.Cell("A1").TryGetValue(out double? _));
        ClassicAssert.IsTrue(ws.Cell("A2").TryGetValue(out double? _));
        ClassicAssert.IsTrue(ws.Cell("A3").TryGetValue(out double? _));
        ClassicAssert.IsFalse(ws.Cell("A4").TryGetValue(out double? _));
    }

    [Test]
    public void CopyRangeAtCellAddress()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");

        ws.Cell("A1")
            .SetValue(2)
            .CellRight()
            .SetValue(3)
            .CellRight()
            .SetValue(5)
            .CellRight()
            .SetValue(7);

        IXLRange range = ws.Range("1:1");

        ws.Cell("B2").CopyFrom(range);

        ClassicAssert.AreEqual(2, ws.Cell("B2").Value);
        ClassicAssert.AreEqual(3, ws.Cell("C2").Value);
        ClassicAssert.AreEqual(5, ws.Cell("D2").Value);
        ClassicAssert.AreEqual(7, ws.Cell("E2").Value);
    }

    [Test]
    public void ValueSetToEmptyString()
    {
        string expected = string.Empty;

        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.Value = new DateTime(2000, 1, 2);
        cell.Value = string.Empty;
        ClassicAssert.AreEqual(expected, cell.GetText());
        ClassicAssert.AreEqual(expected, cell.Value);

        cell.Value = new DateTime(2000, 1, 2);
        cell.SetValue(string.Empty);
        ClassicAssert.AreEqual(expected, cell.GetText());
        ClassicAssert.AreEqual(expected, cell.Value);
    }

    [Test]
    public void ValueSetDateWithShortUserDateFormat()
    {
        // For this test to make sense, user's local date format should be dd/MM/yy (note without the 2 century digits)
        // What happened previously was that the century digits got lost in .ToString() conversion and wrong century was sometimes returned.
        CultureInfo ci = new(CultureInfo.InvariantCulture.LCID);
        ci.DateTimeFormat.ShortDatePattern = "dd/MM/yy";
        Thread.CurrentThread.CurrentCulture = ci;
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        DateTime expected = DateTime.Today.AddYears(20);
        cell.Value = expected;
        DateTime actual = (DateTime)cell.Value;
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void SetStringValueTooLong()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = new DateTime(2018, 5, 15);

            ws.FirstCell().SetValue(new string('A', 32767));

            ClassicAssert.Throws<ArgumentOutOfRangeException>(() =>
                ws.FirstCell().Value = new string('A', 32768)
            );
            ClassicAssert.Throws<ArgumentOutOfRangeException>(() =>
                ws.FirstCell().SetValue(new string('A', 32768))
            );
        }
    }

    [Test]
    public void SetCellValueWipesFormulas()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().FormulaA1 = "=TODAY()";
            ws.FirstCell().Value = "hello world";
            ClassicAssert.IsFalse(ws.FirstCell().HasFormula);

            ws.FirstCell().FormulaA1 = "=TODAY()";
            ws.FirstCell().SetValue("hello world");
            ClassicAssert.IsFalse(ws.FirstCell().HasFormula);
        }
    }

    [Test]
    public void CellValueLineWrapping()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = "hello world";
            ClassicAssert.IsFalse(ws.FirstCell().Style.Alignment.WrapText);

            ws.FirstCell().Value = "hello\r\nworld";
            ClassicAssert.IsTrue(ws.FirstCell().Style.Alignment.WrapText);

            ws.FirstCell().Style.Alignment.WrapText = false;

            ws.FirstCell().SetValue("hello world");
            ClassicAssert.IsFalse(ws.FirstCell().Style.Alignment.WrapText);

            ws.FirstCell().SetValue("hello\r\nworld");
            ClassicAssert.IsTrue(ws.FirstCell().Style.Alignment.WrapText);
        }
    }

    [Test]
    public void TestInvalidXmlCharacters()
    {
        byte[] data;

        using (MemoryStream stream = new())
        {
            XLWorkbook wb = new();
            wb.AddWorksheet("Sheet1").FirstCell().SetValue("\u0018");
            wb.SaveAs(stream);
            data = stream.ToArray();
        }

        using (MemoryStream stream = new(data))
        {
            XLWorkbook wb = new(stream);
            ClassicAssert.AreEqual("\u0018", wb.Worksheets.First().FirstCell().Value);
        }
    }

    [Test]
    public void CanClearDateTimeCellValue()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.AddWorksheet("Sheet1");
                IXLCell c = ws.FirstCell();
                c.SetValue(new DateTime(2017, 10, 08));
                ClassicAssert.AreEqual(XLDataType.DateTime, c.DataType);
                ClassicAssert.AreEqual(new DateTime(2017, 10, 08), c.Value);

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLCell c = ws.FirstCell();
                ClassicAssert.AreEqual(XLDataType.DateTime, c.DataType);
                ClassicAssert.AreEqual(new DateTime(2017, 10, 08), c.Value);

                c.Clear();
                wb.Save();
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLCell c = ws.FirstCell();
                ClassicAssert.AreEqual(XLDataType.Blank, c.DataType);
                ClassicAssert.True(c.IsEmpty());
            }
        }
    }

    [Test]
    public void ClearCellRemovesSparkline()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.SparklineGroups.Add("B1:B3", "C1:E3");

        ws.Cell("B1").Clear();
        ws.Cell("B2").Clear(XLClearOptions.Sparklines);

        ClassicAssert.AreEqual(1, ws.SparklineGroups.Single().Count());
        ClassicAssert.IsFalse(ws.Cell("B1").HasSparkline);
        ClassicAssert.IsFalse(ws.Cell("B2").HasSparkline);
        ClassicAssert.IsTrue(ws.Cell("B3").HasSparkline);
    }

    [Test]
    public void CurrentRegion()
    {
        // Partially based on sample in https://github.com/XlsxSharp/XlsxSharp/issues/120
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.Cell("B1").SetValue("x").CellBelow().SetValue("x").CellBelow().SetValue("x");

            ws.Cell("C1").SetValue("x").CellBelow().SetValue("x").CellBelow().SetValue("x");

            //Deliberately D2
            ws.Cell("D2").SetValue("x").CellBelow().SetValue("x");

            ws.Cell("G1")
                .SetValue("x")
                .CellBelow() // skip a cell
                .CellBelow()
                .SetValue("x")
                .CellBelow()
                .SetValue("x");

            // Deliberately H2
            ws.Cell("H2").SetValue("x").CellBelow().SetValue("x").CellBelow().SetValue("x");

            // A diagonal
            ws.Cell("E8")
                .SetValue("x")
                .CellBelow()
                .CellRight()
                .SetValue("x")
                .CellBelow()
                .CellRight()
                .SetValue("x")
                .CellBelow()
                .CellRight()
                .SetValue("x")
                .CellBelow()
                .CellRight()
                .SetValue("x");

            ClassicAssert.AreEqual("A10:A10", ws.Cell("A10").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("B5:B5", ws.Cell("B5").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("P1:P1", ws.Cell("P1").CurrentRegion.RangeAddress.ToString());

            ClassicAssert.AreEqual("B1:D3", ws.Cell("D3").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("B1:D4", ws.Cell("D4").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("B1:E4", ws.Cell("E4").CurrentRegion.RangeAddress.ToString());

            foreach (IXLCell c in ws.Range("B1:D3").Cells())
            {
                ClassicAssert.AreEqual("B1:D3", c.CurrentRegion.RangeAddress.ToString());
            }

            foreach (IXLCell c in ws.Range("A1:A3").Cells())
            {
                ClassicAssert.AreEqual("A1:D3", c.CurrentRegion.RangeAddress.ToString());
            }

            ClassicAssert.AreEqual("A1:D4", ws.Cell("A4").CurrentRegion.RangeAddress.ToString());

            foreach (IXLCell c in ws.Range("E1:E3").Cells())
            {
                ClassicAssert.AreEqual("B1:E3", c.CurrentRegion.RangeAddress.ToString());
            }

            ClassicAssert.AreEqual("B1:E4", ws.Cell("E4").CurrentRegion.RangeAddress.ToString());

            //// SECOND REGION
            foreach (IXLCell c in ws.Range("F1:F4").Cells())
            {
                ClassicAssert.AreEqual("F1:H4", c.CurrentRegion.RangeAddress.ToString());
            }

            ClassicAssert.AreEqual("F1:H5", ws.Cell("F5").CurrentRegion.RangeAddress.ToString());

            //// DIAGONAL
            ClassicAssert.AreEqual("E8:I12", ws.Cell("E8").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:I12", ws.Cell("F9").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:I12", ws.Cell("G10").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:I12", ws.Cell("H11").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:I12", ws.Cell("I12").CurrentRegion.RangeAddress.ToString());

            ClassicAssert.AreEqual("E8:I12", ws.Cell("G9").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:I12", ws.Cell("F10").CurrentRegion.RangeAddress.ToString());

            ClassicAssert.AreEqual("D7:I12", ws.Cell("D7").CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual("E8:J13", ws.Cell("J13").CurrentRegion.RangeAddress.ToString());

            // Four corners of a sheet
            ClassicAssert.AreEqual("A1:D3", ws.Cell(1, 1).CurrentRegion.RangeAddress.ToString());
            ClassicAssert.AreEqual(
                "XFD1:XFD1",
                ws.Cell(1, XLHelper.MaxColumnNumber).CurrentRegion.RangeAddress.ToString()
            );
            ClassicAssert.AreEqual(
                "XFD1048576:XFD1048576",
                ws.Cell(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber)
                    .CurrentRegion.RangeAddress.ToString()
            );
            ClassicAssert.AreEqual(
                "A1048576:A1048576",
                ws.Cell(XLHelper.MaxRowNumber, 1).CurrentRegion.RangeAddress.ToString()
            );
        }
    }

    // https://github.com/XlsxSharp/XlsxSharp/issues/630
    [Test]
    public void ConsiderEmptyValueAsNumericInSumFormula()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue("Empty");
            ws.Cell("A2").SetValue("Numeric");
            ws.Cell("A3").SetValue("Copy of numeric");

            ws.Cell("B2").SetFormulaA1("=B1");
            ws.Cell("B3").SetFormulaA1("=B2");

            ws.Cell("C2").SetFormulaA1("=SUM(C1)");
            ws.Cell("C3").SetFormulaA1("=C2");

            XLCellValue b1 = ws.Cell("B1").Value;
            XLCellValue b2 = ws.Cell("B2").Value;
            XLCellValue b3 = ws.Cell("B3").Value;

            ClassicAssert.AreEqual(Blank.Value, b1);
            ClassicAssert.AreEqual(0, b2);
            ClassicAssert.AreEqual(0, b3);

            XLCellValue c1 = ws.Cell("C1").Value;
            XLCellValue c2 = ws.Cell("C2").Value;
            XLCellValue c3 = ws.Cell("C3").Value;

            ClassicAssert.AreEqual(Blank.Value, c1);
            ClassicAssert.AreEqual(0, c2);
            ClassicAssert.AreEqual(0, c3);
        }
    }

    [Test]
    public void SetFormulaA1AffectsR1C1()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.FormulaR1C1 = "R[1]C";

            cell.FormulaA1 = "B2";

            ClassicAssert.AreEqual("R[1]C[1]", cell.FormulaR1C1);
        }
    }

    [Test]
    public void SetFormulaR1C1AffectsA1()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.FormulaA1 = "A2";

            cell.FormulaR1C1 = "R[1]C[1]";

            ClassicAssert.AreEqual("B2", cell.FormulaA1);
        }
    }

    [Test]
    [Arguments(" = 1 + SUM({ 1; 7})  - A8  ", "1 + SUM({ 1; 7})  - A8")]
    public void FormulaA1SetterTrimsAndRemovesEqualIfPresent(string formula, string expectedResult)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = formula;
        ClassicAssert.AreEqual(expectedResult, ws.Cell("A1").FormulaA1);
    }

    [Test]
    [Arguments(" =  1 +   R[1]C[7]  ", "1 +   R[1]C[7]")]
    public void FormulaR1C1SetterTrimsAndRemovesEqualIfPresent(
        string formula,
        string expectedResult
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaR1C1 = formula;
        ClassicAssert.AreEqual(expectedResult, ws.Cell("A1").FormulaR1C1);
    }

    [Test]
    public void FormulaWithCircularReferenceFails()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell A1 = ws.Cell("A1");
            IXLCell A2 = ws.Cell("A2");
            A1.FormulaA1 = "A2 + 1";
            A2.FormulaA1 = "A1 + 1";

            InvalidOperationException ex1 = ClassicAssert.Throws<InvalidOperationException>(() =>
                _ = A1.Value
            );
            StringAssert.Contains("cycle", ex1.Message);
            InvalidOperationException ex2 = ClassicAssert.Throws<InvalidOperationException>(() =>
                _ = A2.Value
            );
            StringAssert.Contains("cycle", ex2.Message);
        }
    }

    [Test]
    public void InvalidFormulaShiftProducesRef()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = 1;
                ws.Cell("B1").Value = 2;
                ws.Cell("B2").FormulaA1 = "=A1+B1";

                ClassicAssert.AreEqual(3, ws.Cell("B2").Value);

                ws.Range("B2").CopyTo(ws.Range("A2"));
                string fA2 = ws.Cell("A2").FormulaA1;

                wb.SaveAs(ms);

                ClassicAssert.AreEqual("#REF!+A1", fA2);
            }

            using (XLWorkbook wb2 = new(ms))
            {
                string fA2 = wb2.Worksheets.First().Cell("A2").FormulaA1;
                ClassicAssert.AreEqual("#REF!+A1", fA2);
            }
        }
    }

    [Test]
    public void FormulaWithCircularReferenceFails2()
    {
        IXLCell cell = new XLWorkbook().Worksheets.Add("Sheet1").FirstCell();
        cell.FormulaA1 = "A1";
        ClassicAssert.Throws<InvalidOperationException>(() =>
        {
            XLCellValue _ = cell.Value;
        });
    }

    [Test]
    public void TryGetValueFormulaEvaluationFailReturnFalse()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell A1 = ws.Cell("A1");
            IXLCell A2 = ws.Cell("A2");
            IXLCell A3 = ws.Cell("A3");
            A1.FormulaA1 = "A2 + 1";
            A2.FormulaA1 = "A1 + 1";

            ClassicAssert.IsFalse(A1.TryGetValue(out string _));
            ClassicAssert.IsFalse(A2.TryGetValue(out string _));
            ClassicAssert.IsTrue(A3.TryGetValue(out string _));
        }
    }

    [Test]
    public void ToStringNoFormatString()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLCell c = ws.FirstCell().CellBelow(2).CellRight(3);

        ClassicAssert.AreEqual("D3", c.ToString());
    }

    [Test]
    [Arguments("D3", "A")]
    [Arguments("YEAR(DATE(2018, 1, 1))", "F")]
    [Arguments("YEAR(DATE(2018, 1, 1))", "f")]
    [Arguments("0000.00", "NF")]
    [Arguments("0000.00", "nf")]
    [Arguments("FFFF0000", "fg")]
    [Arguments("Color Theme: Accent5, Tint: 0", "BG")]
    [Arguments("2018.00", "v")]
    public void ToStringFormatString(string expected, string format)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLCell c = ws.FirstCell().CellBelow(2).CellRight(3);

        string formula = "YEAR(DATE(2018, 1, 1))";
        c.FormulaA1 = formula;

        string numberFormat = "0000.00";
        c.Style.NumberFormat.Format = numberFormat;

        c.Style.Font.FontColor = XLColor.Red;
        c.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);

        ClassicAssert.AreEqual(expected, c.ToString(format));

        ClassicAssert.Throws<FormatException>(() => c.ToString("dummy"));
    }

    [Test]
    public void ToStringInvalidFormat()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLCell c = ws.FirstCell();

        ClassicAssert.Throws<FormatException>(() => c.ToString("dummy"));
    }

    [Test]
    public void PropertyActiveIsTrueWhenCellHasSameAddressAsActiveCellInWorksheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.IsNull(ws.ActiveCell);
        ClassicAssert.False(ws.Cell(1, 1).Active);

        ws.ActiveCell = ws.Cell("C4");
        ClassicAssert.True(ws.Cell("C4").Active);
        ClassicAssert.False(ws.Cell("C5").Active);

        ws.ActiveCell = null;
        ClassicAssert.False(ws.Cell("C4").Active);
    }

    [Test]
    public void PropertyActiveDeactivatesCellOnlyWhenTheCellIsActive()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.ActiveCell = ws.Cell("A2");

        ws.Cell("B2").Active = false;
        ClassicAssert.AreEqual(ws.Cell("A2"), ws.ActiveCell);

        ws.Cell("A2").Active = false;
        ClassicAssert.IsNull(ws.ActiveCell);
    }

    [Test]
    public void PropertyActiveSetsCellAsActiveCellOfWorksheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.IsNull(ws.ActiveCell);

        ws.Cell("B2").Active = true;
        ClassicAssert.AreEqual(ws.Cell("B2"), ws.ActiveCell);
    }

    [Test]
    [Arguments("PY(4)", "_xlfn._xlws.PY(4)")]
    [Arguments("5 + py(abs(4) )", "5 + _xlfn._xlws.PY(abs(4) )")]
    [Arguments("COT(COTH(A5 + 2 * SIN(B7)))", "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))")]
    [Arguments(
        "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))",
        "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))"
    )]
    public void FormulaA1AddsPrefixToFutureFunctions(string formula, string expected)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell cell = ws.Cell("A1");
        cell.FormulaA1 = formula;

        ClassicAssert.AreEqual(expected, cell.FormulaA1);
    }

    [Test]
    [Arguments("PY(4)", "_xlfn._xlws.PY(4)")]
    [Arguments("5 + py(abs(4) )", "5 + _xlfn._xlws.PY(abs(4) )")]
    [Arguments(
        "COT(COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))",
        "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))"
    )]
    [Arguments(
        "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))",
        "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))"
    )]
    public void FormulaR1C1AddsPrefixToFutureFunctions(string formula, string expected)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell cell = ws.Cell("A1");
        cell.FormulaR1C1 = formula;

        ClassicAssert.AreEqual(expected, cell.FormulaR1C1);
    }

    [Test]
    public void FormulaA1AddsPrefixToAllFutureFunctions()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell cell = ws.Cell("A1");
        foreach ((string simpleName, string prefixedName) in XLConstants.FutureFunctionMap.Value)
        {
            cell.FormulaA1 = simpleName + "()";
            ClassicAssert.AreEqual(prefixedName + "()", cell.FormulaA1);
        }
    }
}
