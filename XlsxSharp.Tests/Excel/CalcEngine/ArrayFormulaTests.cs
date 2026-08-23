using System;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class ArrayFormulaTests
{
    [Test]
    public void ArrayFormulaIsSaved() =>
        TestHelper.CreateAndCompare(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();
                ws.Range("A1:B2").FormulaArrayA1 = "1+2";
            },
            @"Other\Formulas\ArrayFormula.xlsx"
        );

    [Test]
    public void ArrayFormulaCanBeLoaded() =>
        TestHelper.LoadAndAssert(
            wb =>
            {
                IXLWorksheet ws = wb.Worksheets.First();

                foreach (IXLCell arrayFormulaCell in ws.Range("A1:B2").Cells())
                {
                    ClassicAssert.AreEqual("1+2", arrayFormulaCell.FormulaA1);
                    ClassicAssert.AreEqual(
                        "A1:B2",
                        arrayFormulaCell.FormulaReference.ToStringRelative()
                    );
                }

                IXLCell outsideCell = ws.Cell("A3");
                ClassicAssert.IsEmpty(outsideCell.FormulaA1);
                ClassicAssert.Null(outsideCell.FormulaReference);
            },
            @"Other\Formulas\ArrayFormula.xlsx"
        );

    [Test]
    public void CanBeOnlyForOneCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell oneCell = ws.Cell("B3");

        oneCell.AsRange().FormulaArrayA1 = "2+5";

        ClassicAssert.True(oneCell.HasArrayFormula);
        ClassicAssert.AreEqual("2+5", oneCell.FormulaA1);
        ClassicAssert.AreEqual("B3:B3", oneCell.FormulaReference.ToStringRelative());
    }

    [Test]
    [Arguments("B2:C3")]
    [Arguments("B2:C4")]
    [Arguments("A1:D7")]
    public void SettingValueToContainingRangeClearsArrayFormula(string containingRange)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange arrayFormulaRange = ws.Range("B2:C3");
        arrayFormulaRange.FormulaArrayA1 = "5";

        ws.Range(containingRange).Value = Blank.Value;

        foreach (IXLCell cell in arrayFormulaRange.Cells())
        {
            ClassicAssert.AreEqual(Blank.Value, cell.Value);
            ClassicAssert.False(cell.HasArrayFormula);
            ClassicAssert.IsEmpty(cell.FormulaA1);
            ClassicAssert.Null(cell.FormulaReference);
        }
    }

    [Test]
    [Arguments("B2:D3")]
    [Arguments("A1:E4")]
    public void SettingFormulaToContainingRangeClearsOriginalArrayFormula(string overlapRange)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("B2:D3").FormulaArrayA1 = "1";

        ClassicAssert.DoesNotThrow(() => ws.Range(overlapRange).FormulaArrayA1 = "2");
    }

    [Test]
    [Arguments("B2:B2")]
    [Arguments("B2:B3")]
    [Arguments("A1:C3")]
    [Arguments("D2:F3")]
    [Arguments("C:C")]
    [Arguments("2:2")]
    public void ArrayFormulaCantPartiallyOverlapWithAnotherArrayFormula(string partialOverlapRange)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("B2:D3").FormulaArrayA1 = "1";

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            ws.Range(partialOverlapRange).FormulaArrayA1 = "2"
        );
        ClassicAssert.AreEqual(
            "Can't create array function that partially covers another array function.",
            ex.Message
        );
    }

    [Test]
    [Arguments("A1:B2")]
    [Arguments("A2")]
    public void ArrayFormulaCantOverlapWithMergedRange(string partialOverlapRange)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Range("A1:A2").Merge();

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            ws.Range(partialOverlapRange).FormulaArrayA1 = "1"
        );
        ClassicAssert.AreEqual("Can't create array function over a merged range.", ex.Message);
    }

    [Test]
    [Arguments("A1:B2")]
    [Arguments("A1:C1")]
    public void ArrayFormulaCantOverlapWithTable(string formulaRange)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Name";
        ws.Cell("A2").Value = 5;
        ws.Range("A1:A2").CreateTable();

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            ws.Range(formulaRange).FormulaArrayA1 = "1"
        );
        ClassicAssert.AreEqual("Can't create array function over a table.", ex.Message);
    }

    [Test]
    public void SettingArrayFormulaInvalidatesCells()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.False(ws.Cell("A1").NeedsRecalculation);
        ClassicAssert.False(ws.Cell("A2").NeedsRecalculation);

        ws.Range("A1:A2").FormulaArrayA1 = "ABS(-3)";

        ClassicAssert.True(ws.Cell("A1").NeedsRecalculation);
        ClassicAssert.True(ws.Cell("A2").NeedsRecalculation);
    }

    [Test]
    public void ReferencingItselfIsCircularError()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "A2";
        ws.Range("A2").FormulaArrayA1 = "A1";

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            _ = ws.Cell("A2").Value
        );
        ClassicAssert.AreEqual("Formula in a cell Sheet1!A1 is part of a cycle.", ex.Message);
    }
}
