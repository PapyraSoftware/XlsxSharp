using XlsxSharp.Excel;
using XlsxSharp.Excel.DataValidation;

namespace XlsxSharp.Tests.Excel.DataValidations;

public class DataValidationShiftTests
{
    [Test]
    public void DataValidationShiftedOnColumnInsert()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).InsertColumnsAfter(2);
        IXLDataValidation[] dv = [.. ws.DataValidations];

        ClassicAssert.AreEqual(5, dv.Length);
        ClassicAssert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("A2:D2", dv[1].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("A3:E3", dv[2].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("B4:D6", dv[3].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("E7:F7", dv[4].Ranges.Single().RangeAddress.ToString());
    }

    [Test]
    public void DataValidationShiftedOnRowInsert()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).InsertRowsBelow(2);
        IXLDataValidation[] dv = [.. ws.DataValidations];

        ClassicAssert.AreEqual(5, dv.Length);
        ClassicAssert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("B1:B4", dv[1].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("C1:C5", dv[2].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("D2:F4", dv[3].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("G6:G7", dv[4].Ranges.Single().RangeAddress.ToString());
    }

    [Test]
    public void DataValidationShiftedOnColumnDelete()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).Delete();
        IXLDataValidation[] dv = [.. ws.DataValidations];

        ClassicAssert.AreEqual(4, dv.Length);
        ClassicAssert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("A2:A2", dv[1].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("A3:B3", dv[2].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("B7:C7", dv[3].Ranges.Single().RangeAddress.ToString());
    }

    [Test]
    public void DataValidationShiftedOnRowDelete()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).Delete();
        IXLDataValidation[] dv = [.. ws.DataValidations];

        ClassicAssert.AreEqual(4, dv.Length);
        ClassicAssert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("B1:B1", dv[1].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("C1:C2", dv[2].Ranges.Single().RangeAddress.ToString());
        ClassicAssert.AreEqual("G3:G4", dv[3].Ranges.Single().RangeAddress.ToString());
    }

    [Test]
    [Arguments(new[] { "A10:A11" }, "1-2", new[] { "A8:A9" })]
    [Arguments(new[] { "A10,A11" }, "1-2", new[] { "A8:A8 A9:A9" })]
    [Arguments(new[] { "A10", "A11" }, "1-2", new[] { "A8:A8", "A9:A9" })]
    public void DataValidationsAreShiftedWhenRowsAboveAreDeleted(
        string[] initialDvs,
        string rowsToDelete,
        string[] shiftedDvs
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        foreach (string initialDv in initialDvs)
        {
            ws.Ranges(initialDv).CreateDataValidation();
        }

        ws.Rows(rowsToDelete).Delete();

        IEnumerable<string> resultDvs = ws.DataValidations.Select(dv => dv.Ranges.ToSpaceList());
        ClassicAssert.AreEqual(shiftedDvs, resultDvs);
    }

    [Test]
    public void DataValidationsIsRemovedWhenItsAreaIsDeleted()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A10").CreateDataValidation();

        ws.Range("A10").Delete(XLShiftDeletedCells.ShiftCellsUp);

        ClassicAssert.IsEmpty(ws.DataValidations);
    }

    [Test]
    public void DataValidationsCanSplitItsAreaWhenInsertedOrDeletedAreaIntersectsItsArea()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A10:C12").CreateDataValidation();

        ws.Range("B12").Delete(XLShiftDeletedCells.ShiftCellsUp);

        ClassicAssert.AreEqual(
            "A10:C11 A12:A12 C12:C12",
            ws.DataValidations.Single().Ranges.ToSpaceList()
        );
    }

    [Test]
    public void DataValidationShiftedTruncateRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("DataValidationShift");
        ws.AsRange().CreateDataValidation().WholeNumber.Between(0, 1);
        IXLDataValidation dv = ws.DataValidations.Single();

        ws.Row(2).InsertRowsAbove(1);
        ClassicAssert.IsTrue(dv.Ranges.Single().RangeAddress.IsValid);
        ClassicAssert.AreEqual(
            $"1:{XLHelper.MaxRowNumber}",
            dv.Ranges.Single().RangeAddress.ToString()
        );

        ws.Column(2).InsertColumnsAfter(1);
        ClassicAssert.IsTrue(dv.Ranges.Single().RangeAddress.IsValid);
        ClassicAssert.AreEqual(
            $"1:{XLHelper.MaxRowNumber}",
            dv.Ranges.Single().RangeAddress.ToString()
        );
    }
}
