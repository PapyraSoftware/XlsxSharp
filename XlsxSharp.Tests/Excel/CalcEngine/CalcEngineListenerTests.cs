using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

/// <summary>
/// Tests that calc engine adjusts its internal state in response to changes of workbook structure.
/// </summary>
internal class CalcEngineListenerTests
{
    [Test]
    public void Formulas_dependent_on_specific_sheet_are_dirty_after_sheet_addition()
    {
        using XLWorkbook wb = new();
        IXLWorksheet sutWs = wb.AddWorksheet();
        sutWs.Cell("A1").FormulaA1 = "new!A1";
        ClassicAssert.AreEqual(XLError.CellReference, sutWs.Cell("A1").Value);

        IXLWorksheet newWs = wb.AddWorksheet("new");
        newWs.Cell("A1").Value = 5;

        // Cell contains last calculated value
        ClassicAssert.AreEqual(XLError.CellReference, sutWs.Cell("A1").CachedValue);

        // But once asked for real value, it calculates it.
        ClassicAssert.True(sutWs.Cell("A1").NeedsRecalculation);
        ClassicAssert.AreEqual(5.0, sutWs.Cell("A1").Value);
    }

    [Test]
    public void Formulas_dependent_on_specific_sheet_are_dirty_after_sheet_deletion()
    {
        using XLWorkbook wb = new();
        IXLWorksheet keptWs = wb.AddWorksheet();
        IXLWorksheet deletedWs = wb.AddWorksheet("deleted");

        deletedWs.Cell("A1").Value = 5;
        keptWs.Cell("A1").FormulaA1 = "deleted!A1";
        ClassicAssert.AreEqual(5.0, keptWs.Cell("A1").Value);

        deletedWs.Delete();

        // Cell contains last calculated value
        ClassicAssert.AreEqual(5.0, keptWs.Cell("A1").CachedValue);

        // But once asked for real value, it calculates it.
        ClassicAssert.True(keptWs.Cell("A1").NeedsRecalculation);
        ClassicAssert.AreEqual(XLError.CellReference, keptWs.Cell("A1").Value);
    }

    [Test]
    public void Formulas_are_shifted_when_area_is_added_and_cells_shifted_down()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "B1*2";
        ws.Cell("B1").FormulaA1 = "C1*2";
        ws.Cell("C1").FormulaA1 = "1+2";

        ws.RecalculateAllFormulas();

        ws.Range("A1:B1").InsertRowsAbove(2);

        ClassicAssert.AreEqual(12.0, ws.Cell("A3").Value);
        ClassicAssert.False(ws.Cell("A3").NeedsRecalculation);
        ClassicAssert.False(ws.Cell("B3").NeedsRecalculation);

        // Dependency tree should pick up the change
        ws.Cell("C1").FormulaA1 = "2+2";
        ClassicAssert.True(ws.Cell("A3").NeedsRecalculation);
        ClassicAssert.True(ws.Cell("B3").NeedsRecalculation);
        ClassicAssert.AreEqual(16.0, ws.Cell("A3").Value);
    }

    [Test]
    public void Formulas_are_shifted_when_area_is_added_and_cells_shifted_right()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "A2*2";
        ws.Cell("A2").FormulaA1 = "A3*2";
        ws.Cell("A3").FormulaA1 = "1+2";

        ws.RecalculateAllFormulas();

        ws.Cell("A2").InsertCellsBefore(4);

        ClassicAssert.AreEqual(12.0, ws.Cell("A1").Value);
        ClassicAssert.False(ws.Cell("E2").NeedsRecalculation);

        // Dependency tree should pick up the change
        ws.Cell("A3").FormulaA1 = "2+2";
        ClassicAssert.True(ws.Cell("E2").NeedsRecalculation);
        ClassicAssert.True(ws.Cell("A1").NeedsRecalculation);
        ClassicAssert.AreEqual(16.0, ws.Cell("A1").Value);
    }

    [Test]
    public void Formulas_are_shifted_when_area_is_deleted_and_cells_shifted_up()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A5").FormulaA1 = "1+2";
        ws.Cell("B5").FormulaA1 = "A5*2";
        ws.Cell("C5").FormulaA1 = "B5*2";

        ws.RecalculateAllFormulas();

        ws.Range("B2:C4").Delete(XLShiftDeletedCells.ShiftCellsUp);

        ClassicAssert.AreEqual(12.0, ws.Cell("C2").Value);
        ClassicAssert.False(ws.Cell("B2").NeedsRecalculation);
        ClassicAssert.False(ws.Cell("A2").NeedsRecalculation);

        // Dependency tree should pick up the change
        ws.Cell("A5").FormulaA1 = "2+2";
        ClassicAssert.True(ws.Cell("B2").NeedsRecalculation);
        ClassicAssert.True(ws.Cell("C2").NeedsRecalculation);
        ClassicAssert.AreEqual(16.0, ws.Cell("C2").Value);
    }

    [Test]
    public void Formulas_are_shifted_when_area_is_deleted_and_cells_shifted_left()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "1+2";
        ws.Cell("E2").FormulaA1 = "D1*2";
        ws.Cell("D3").FormulaA1 = "E2*2";

        ws.RecalculateAllFormulas();

        ws.Range("A1:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);

        ClassicAssert.AreEqual(12.0, ws.Cell("A3").Value);
        ClassicAssert.False(ws.Cell("B2").NeedsRecalculation);
        ClassicAssert.False(ws.Cell("A1").NeedsRecalculation);

        // Dependency tree should pick up the change
        ws.Cell("A1").FormulaA1 = "2+2";
        ClassicAssert.True(ws.Cell("B2").NeedsRecalculation);
        ClassicAssert.True(ws.Cell("A3").NeedsRecalculation);
        ClassicAssert.AreEqual(16.0, ws.Cell("A3").Value);
    }
}
