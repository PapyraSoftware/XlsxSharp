using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Cells;

public class XlCellFormulaTests
{
    [Test]
    public void CellFormulaIsStrippedOfEqualSign()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 1).FormulaA1 = "=B1";
        ClassicAssert.AreEqual("B1", ws.Cell(1, 1).FormulaA1);
    }

    [Test]
    public void DataTableMaintainProperties() =>
        TestHelper.LoadSaveAndCompare(
            @"Other\Formulas\DataTableFormula-Excel-Input.xlsx",
            @"Other\Formulas\DataTableFormula-Output.xlsx"
        );
}
