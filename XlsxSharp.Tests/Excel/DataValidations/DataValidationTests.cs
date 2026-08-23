using System;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Tests.Excel.DataValidations;

public class DataValidationTests
{
    [Test]
    public void ValidationReferenceListValuesFromSeparateSheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet valuesSheet = wb.Worksheets.Add("ValuesSheet");
        IXLCell cell = valuesSheet.Cell("E1");
        cell.SetValue("Value 1");
        cell = cell.CellBelow();
        cell.SetValue("Value 2");
        cell = cell.CellBelow();
        cell.SetValue("Value 3");
        cell = cell.CellBelow();
        cell.SetValue("Value 4");

        IXLWorksheet uiSheet = wb.Worksheets.Add("UI Sheet");
        uiSheet
            .Cell("A1")
            .SetValue("Cell below has validation with references to the 'ValuesSheet'.");
        cell = uiSheet.Cell("A2");
        cell.GetDataValidation().List(valuesSheet.Range("ValuesSheet!$E$1:$E$4"));

        ClassicAssert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
        ClassicAssert.AreEqual("ValuesSheet!$E$1:$E$4", cell.GetDataValidation().Value);
    }

    [Test]
    public void Validation1()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Data Validation Issue");
        IXLCell cell = ws.Cell("E1");
        cell.SetValue("Value 1");
        cell = cell.CellBelow();
        cell.SetValue("Value 2");
        cell = cell.CellBelow();
        cell.SetValue("Value 3");
        cell = cell.CellBelow();
        cell.SetValue("Value 4");

        ws.Cell("A1").SetValue("Cell below has Validation Only.");
        cell = ws.Cell("A2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));

        ws.Cell("B1").SetValue("Cell below has Validation with a title.");
        cell = ws.Cell("B2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
        cell.GetDataValidation().InputTitle = "Title for B2";

        ClassicAssert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
        ClassicAssert.AreEqual("'Data Validation Issue'!$E$1:$E$4", cell.GetDataValidation().Value);
        ClassicAssert.AreEqual("Title for B2", cell.GetDataValidation().InputTitle);

        ws.Cell("C1").SetValue("Cell below has Validation with a message.");
        cell = ws.Cell("C2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
        cell.GetDataValidation().InputMessage = "Message for C2";

        ClassicAssert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
        ClassicAssert.AreEqual("'Data Validation Issue'!$E$1:$E$4", cell.GetDataValidation().Value);
        ClassicAssert.AreEqual("Message for C2", cell.GetDataValidation().InputMessage);

        ws.Cell("D1").SetValue("Cell below has Validation with title and message.");
        cell = ws.Cell("D2");
        cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
        cell.GetDataValidation().InputTitle = "Title for D2";
        cell.GetDataValidation().InputMessage = "Message for D2";

        ClassicAssert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
        ClassicAssert.AreEqual("'Data Validation Issue'!$E$1:$E$4", cell.GetDataValidation().Value);
        ClassicAssert.AreEqual("Title for D2", cell.GetDataValidation().InputTitle);
        ClassicAssert.AreEqual("Message for D2", cell.GetDataValidation().InputMessage);
    }

    [Test]
    public void Validation2()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("Sheet1!A1");

        IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");
        ws2.Cell("A1").SetValue("B");
        ws.Cell("B1").CopyTo(ws2.Cell("B1"));

        ClassicAssert.AreEqual("Sheet1!A1", ws2.Cell("B1").GetDataValidation().Value);
    }

    [Test, Skip("Wait for proper formula shifting (#686)")]
    public void Validation3()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.FirstRow().InsertRowsAbove(1);

        ClassicAssert.AreEqual("A2", ws.Cell("B2").GetDataValidation().Value);
    }

    [Test]
    public void Validation4()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.Cell("B1").CopyTo(ws.Cell("B2"));
        ClassicAssert.AreEqual("A2", ws.Cell("B2").GetDataValidation().Value);
    }

    [Test, Skip("Wait for proper formula shifting (#686)")]
    public void Validation5()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.FirstColumn().InsertColumnsBefore(1);

        ClassicAssert.AreEqual("B1", ws.Cell("C1").GetDataValidation().Value);
    }

    [Test]
    public void Validation6()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").SetValue("A");
        ws.Cell("B1").CreateDataValidation().Custom("A1");
        ws.Cell("B1").CopyTo(ws.Cell("C1"));
        ClassicAssert.AreEqual("B1", ws.Cell("C1").GetDataValidation().Value);
    }

    [Test]
    public void ValidationPersistsOnCellDataValidation()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("People");

        ws.FirstCell().SetValue("Categories").CellBelow().SetValue("A").CellBelow().SetValue("B");

        IXLTable table = ws.RangeUsed().CreateTable();

        IXLDataValidation dv = table.DataRange.CreateDataValidation();
        dv.ErrorTitle = "Error";

        ClassicAssert.AreEqual("Error", table.DataRange.FirstCell().GetDataValidation().ErrorTitle);
    }

    [Test]
    public void ValidationPersistsOnWorksheetDataValidations()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("People");

        ws.FirstCell().SetValue("Categories").CellBelow().SetValue("A");

        IXLTable table = ws.RangeUsed().CreateTable();

        IXLDataValidation dv = table.DataRange.CreateDataValidation();
        dv.ErrorTitle = "Error";

        ClassicAssert.AreEqual("Error", ws.DataValidations.Single().ErrorTitle);
    }

    [Test]
    [Arguments("A1:C3", 5, false, "A1:C3")]
    [Arguments("A1:C3", 2, false, "A1:C4")]
    [Arguments("A1:C3", 1, false, "A2:C4")]
    [Arguments("A1:C3", 5, true, "A1:C3")]
    [Arguments("A1:C3", 2, true, "A1:C4")]
    [Arguments("A1:C3", 1, true, "A2:C4")]
    public void DataValidationShiftedOnRowInsert(
        string initialAddress,
        int rowNum,
        bool setValue,
        string expectedAddress
    )
    {
        //Arrange
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("DataValidation");
        IXLDataValidation validation = ws.Range(initialAddress).CreateDataValidation();
        validation.WholeNumber.Between(0, 100);
        if (setValue)
        {
            ws.Range(initialAddress).Value = 50;
        }

        //Act
        ws.Row(rowNum).InsertRowsAbove(1);

        //Assert
        ClassicAssert.AreEqual(1, ws.DataValidations.Count());
        ClassicAssert.AreEqual(1, ws.DataValidations.First().Ranges.Count());
        ClassicAssert.AreEqual(
            expectedAddress,
            ws.DataValidations.First().Ranges.First().RangeAddress.ToString()
        );
    }

    [Test]
    [Arguments("A1:C3", 5, false, "A1:C3")]
    [Arguments("A1:C3", 2, false, "A1:D3")]
    [Arguments("A1:C3", 1, false, "B1:D3")]
    [Arguments("A1:C3", 5, true, "A1:C3")]
    [Arguments("A1:C3", 2, true, "A1:D3")]
    [Arguments("A1:C3", 1, true, "B1:D3")]
    public void DataValidationShiftedOnColumnInsert(
        string initialAddress,
        int columnNum,
        bool setValue,
        string expectedAddress
    )
    {
        //Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("DataValidation");
        IXLDataValidation validation = ws.Range(initialAddress).CreateDataValidation();
        validation.WholeNumber.Between(0, 100);
        if (setValue)
        {
            ws.Range(initialAddress).Value = 50;
        }

        //Act
        ws.Column(columnNum).InsertColumnsBefore(1);

        //Assert
        ClassicAssert.AreEqual(1, ws.DataValidations.Count());
        ClassicAssert.AreEqual(1, ws.DataValidations.First().Ranges.Count());
        ClassicAssert.AreEqual(
            expectedAddress,
            ws.DataValidations.First().Ranges.First().RangeAddress.ToString()
        );
    }

    [Test]
    public void DataValidationClearSplitsRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("DataValidation");
            IXLDataValidation validation = ws.Range("A1:C3").CreateDataValidation();
            validation.WholeNumber.Between(0, 100);

            //Act
            ws.Cell("B2").Clear(XLClearOptions.DataValidation);

            //Assert
            ClassicAssert.IsFalse(ws.Cell("B2").HasDataValidation);
            ClassicAssert.IsTrue(
                ws.Range("A1:C3")
                    .Cells()
                    .Where(c => c.Address.ToString() != "B2")
                    .All(c => c.HasDataValidation)
            );
        }
    }

    [Test]
    public void NewDataValidationSplitsRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("DataValidation");
            IXLDataValidation validation = ws.Range("A1:C3").CreateDataValidation();
            validation.WholeNumber.Between(10, 100);

            //Act
            ws.Cell("B2").CreateDataValidation().WholeNumber.Between(-100, -0);

            //Assert
            ClassicAssert.AreEqual("-100", ws.Cell("B2").GetDataValidation().MinValue);
            ClassicAssert.IsTrue(
                ws.Range("A1:C3")
                    .Cells()
                    .Where(c => c.Address.ToString() != "B2")
                    .All(c => c.HasDataValidation)
            );
            ClassicAssert.IsTrue(
                ws.Range("A1:C3")
                    .Cells()
                    .Where(c => c.Address.ToString() != "B2")
                    .All(c => c.GetDataValidation().MinValue == "10")
            );
        }
    }

    [Test]
    public void ListLengthOverflow()
    {
        string values = string.Join(
            ",",
            Enumerable.Range(1, 20).Select(i => Guid.NewGuid().ToString("N"))
        );

        ClassicAssert.True(values.Length > 255);

        using (XLWorkbook wb = new())
        {
            IXLDataValidation dv = wb.AddWorksheet("Sheet 1").Cell(1, 1).GetDataValidation();

            ClassicAssert.Throws<ArgumentOutOfRangeException>(() => dv.List(values));
            ClassicAssert.Throws<ArgumentOutOfRangeException>(() =>
            {
                dv.TextLength.Between(0, 5);
                dv.MinValue = values;
            });

            ClassicAssert.Throws<ArgumentOutOfRangeException>(() =>
            {
                dv.TextLength.Between(0, 5);
                dv.MaxValue = values;
            });
        }
    }

    [Test]
    public void DataValidationHasWorksheetAndRangesWhenCreated()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLRange range = ws.Range("A1:A3");

            IXLDataValidation dv = range.CreateDataValidation();

            ClassicAssert.AreSame(ws, ((XLDataValidation)dv).Worksheet);
            ClassicAssert.AreSame(range, dv.Ranges.Single());
        }
    }

    [Test]
    public void CanAddRangeFromSameWorksheet()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLRange range1 = ws.Range("A1:A3");
            IXLRange range2 = ws.Range("C1:C3");
            IXLRanges ranges3 = ws.Ranges("D1:D3,F1:F3");
            IXLDataValidation dv = range1.CreateDataValidation();

            dv.AddRange(range2);
            dv.AddRanges(ranges3);

            ClassicAssert.IsTrue(dv.Ranges.Any(r => r == range1));
            ClassicAssert.IsTrue(dv.Ranges.Any(r => r == range2));
            ClassicAssert.IsTrue(dv.Ranges.Any(r => r == ranges3.First()));
            ClassicAssert.IsTrue(dv.Ranges.Any(r => r == ranges3.Last()));
        }
    }

    [Test]
    public void CanAddRangeFromAnotherWorksheet()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet();
            IXLWorksheet ws2 = wb.AddWorksheet();
            IXLRange range1 = ws1.Range("A1:A3");
            IXLRange range2 = ws2.Range("C1:C3");
            IXLDataValidation dv = range1.CreateDataValidation();

            dv.AddRange(range2);

            ClassicAssert.IsTrue(
                dv.Ranges.Any(r =>
                    r != range2 && r.RangeAddress.ToString() == range2.RangeAddress.ToString()
                )
            );
        }
    }

    [Test]
    public void CanClearRanges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLRange range1 = ws.Range("A1:A3");
            IXLRange range2 = ws.Range("C1:C3");
            IXLRanges ranges3 = ws.Ranges("D1:D3,F1:F3");
            IXLDataValidation dv = range1.CreateDataValidation();
            dv.AddRange(range2);
            dv.AddRanges(ranges3);

            dv.ClearRanges();

            ClassicAssert.IsEmpty(dv.Ranges);
        }
    }

    [Test]
    public void CanRemoveExistingRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLRange range1 = ws.Range("A1:A3");
            IXLRange range2 = ws.Range("C1:C3");

            IXLDataValidation dv = range1.CreateDataValidation();
            dv.AddRange(range2);

            dv.RemoveRange(range1);

            ClassicAssert.AreSame(range2, dv.Ranges.Single());
        }
    }

    [Test]
    public void RemovingExistingRangeDoesNoFail()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLRange range1 = ws.Range("A1:A3");
            IXLRange range2 = ws.Range("C1:C3");

            IXLDataValidation dv = range1.CreateDataValidation();

            dv.RemoveRange(range2);
            dv.RemoveRange(null);

            ClassicAssert.AreSame(range1, dv.Ranges.Single());
        }
    }
}
