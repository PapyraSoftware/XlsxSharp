using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class FormulaCachingTests
{
    [Test]
    public void StaticCellDoesNotNeedRecalculation()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell cell = sheet.Cell(1, 1);
            cell.Value = "1234567";

            ClassicAssert.IsFalse(cell.NeedsRecalculation);
        }
    }

    [Test]
    public void EditCellInvalidatesDependentCells()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell cell = sheet.Cell(1, 1);
            IXLCell dependentCell = sheet.Cell(2, 1);
            dependentCell.FormulaA1 = "=A1";
            XLCellValue _ = dependentCell.Value;

            cell.Value = "1234567";

            ClassicAssert.IsTrue(dependentCell.NeedsRecalculation);
        }
    }

    [Test]
    public void EditFormulaA1InvalidatesDependentCells()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a1 = sheet.Cell("A1");
            IXLCell a2 = sheet.Cell("A2");
            IXLCell a3 = sheet.Cell("A3");
            IXLCell a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            XLCellValue res1 = a4.Value;
            a2.FormulaA1 = "=A1*20";
            XLCellValue res2 = a4.Value;

            ClassicAssert.AreEqual(15 + 150 + 1500, res1);
            ClassicAssert.AreEqual(15 + 300 + 3000, res2);
        }
    }

    [Test]
    public void EditFormulaR1C1InvalidatesDependentCells()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a1 = sheet.Cell("A1");
            IXLCell a2 = sheet.Cell("A2");
            IXLCell a3 = sheet.Cell("A3");
            IXLCell a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            XLCellValue res1 = a4.Value;
            a2.FormulaR1C1 = "=R[-1]C*2";
            XLCellValue res2 = a4.Value;

            ClassicAssert.AreEqual(15 + 150 + 1500, res1);
            ClassicAssert.AreEqual(15 + 30 + 300, res2);
        }
    }

    [Test]
    public void InsertRowInvalidatesValues()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a4 = sheet.Cell("A4");
            a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

            ClassicAssert.AreEqual(3, a4.Value);

            sheet.Row(2).InsertRowsAbove(2);

            ClassicAssert.AreEqual(5, sheet.Cell("A6").Value);
        }
    }

    [Test]
    public void DeleteRowModifiesFormulaAndInvalidatesValues()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell original = sheet.Cell("A4");
            original.FormulaA1 = "=COUNTBLANK(A1:A3)";

            ClassicAssert.AreEqual(3, original.Value);

            sheet.Row(2).Delete();

            IXLCell shifted = sheet.Cell("A3");
            ClassicAssert.AreEqual("COUNTBLANK(A1:A2)", shifted.FormulaA1);
            ClassicAssert.AreEqual(2, shifted.Value);
        }
    }

    [Test]
    public void ChainedCalculationPreservesIntermediateValues()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a1 = sheet.Cell("A1");
            IXLCell a2 = sheet.Cell("A2");
            IXLCell a3 = sheet.Cell("A3");
            IXLCell a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";

            a1.Value = 15;
            XLCellValue res = a4.Value;

            ClassicAssert.AreEqual(15 + 150 + 1500, res);
            ClassicAssert.IsFalse(a4.NeedsRecalculation);
            ClassicAssert.IsFalse(a3.NeedsRecalculation);
            ClassicAssert.IsFalse(a2.NeedsRecalculation);
            ClassicAssert.AreEqual(150, a2.CachedValue);
            ClassicAssert.AreEqual(1500, a3.CachedValue);
            ClassicAssert.AreEqual(15 + 150 + 1500, a4.CachedValue);
        }
    }

    [Test]
    public void EditingAffectsDependentCells()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a1 = sheet.Cell("A1");
            IXLCell a2 = sheet.Cell("A2");
            IXLCell a3 = sheet.Cell("A3");
            IXLCell a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            XLCellValue res1 = a4.Value;
            a1.Value = 20;
            XLCellValue res2 = a4.Value;

            ClassicAssert.AreEqual(15 + 150 + 1500, res1);
            ClassicAssert.AreEqual(20 + 200 + 2000, res2);
        }
    }

    [Test]
    [Arguments("C4", new string[] { "C5" })]
    [Arguments("D4", new string[] { })]
    [Arguments("A1", new string[] { "A2", "A3", "A4", "C1", "C2", "C3", "C5" })]
    [Arguments("B2", new string[] { "B3", "B4", "C2", "C3", "C5" })]
    [Arguments("C2", new string[] { "C5" })]
    public void EditingDoesNotAffectNonDependingCells(string changedCell, string[] affectedCells)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            sheet.Cell("A2").FormulaA1 = "A1+1";
            sheet.Cell("A3").FormulaA1 = "SUM(A1:A2)";
            sheet.Cell("A4").FormulaA1 = "SUM(A1:A3)";
            sheet.Cell("B2").FormulaA1 = "B1+1";
            sheet.Cell("B3").FormulaA1 = "SUM(B1:B2)";
            sheet.Cell("B4").FormulaA1 = "SUM(B1:B3)";
            sheet.Cell("C1").FormulaA1 = "SUM(A1:B1)";
            sheet.Cell("C2").FormulaA1 = "SUM(A2:B2)";
            sheet.Cell("C3").FormulaA1 = "SUM(A3:B3)";
            sheet.Cell("C5").FormulaA1 = "SUM($A$1:$C$4)";
            sheet.RecalculateAllFormulas();
            IXLCells allCells = sheet.CellsUsed();

            sheet.Cell(changedCell).Value = 100;
            IEnumerable<IXLCell> modifiedCells = allCells.Where(cell => cell.NeedsRecalculation);

            ClassicAssert.AreEqual(affectedCells.Length, modifiedCells.Count());
            foreach (string cellAddress in affectedCells)
            {
                ClassicAssert.IsTrue(
                    modifiedCells.Any(cell => cell.Address.ToString() == cellAddress),
                    string.Format(
                        "Cell {0} is expected to need recalculation, but it does not",
                        cellAddress
                    )
                );
            }
        }
    }

    [Test]
    public void CircularReferenceFailsCalculating()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a1 = sheet.Cell("A1");
            IXLCell a2 = sheet.Cell("A2");
            IXLCell a3 = sheet.Cell("A3");
            IXLCell a4 = sheet.Cell("A4");

            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=A3*10";
            a1.FormulaA1 = "A2+A3+A4";

            TestDelegate getValueA1 = new(() => _ = a1.Value);
            TestDelegate getValueA2 = new(() => _ = a2.Value);
            TestDelegate getValueA3 = new(() => _ = a3.Value);
            TestDelegate getValueA4 = new(() => _ = a4.Value);

            ClassicAssert.Throws(typeof(InvalidOperationException), getValueA1);
            ClassicAssert.Throws(typeof(InvalidOperationException), getValueA2);
            ClassicAssert.Throws(typeof(InvalidOperationException), getValueA3);
            ClassicAssert.Throws(typeof(InvalidOperationException), getValueA4);
        }
    }

    [Test]
    public void CircularReferenceRecalculationNeededDoesNotFail()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet = wb.Worksheets.Add("TestSheet");
            IXLCell a1 = sheet.Cell("A1");
            IXLCell a2 = sheet.Cell("A2");
            IXLCell a3 = sheet.Cell("A3");
            IXLCell a4 = sheet.Cell("A4");

            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=A3*10";
            XLCellValue _ = a4.Value;
            a1.FormulaA1 = "=SUM(A2:A4)";

            bool recalcNeededA1 = a1.NeedsRecalculation;
            bool recalcNeededA2 = a2.NeedsRecalculation;
            bool recalcNeededA3 = a3.NeedsRecalculation;
            bool recalcNeededA4 = a4.NeedsRecalculation;

            ClassicAssert.IsTrue(recalcNeededA1);
            ClassicAssert.IsTrue(recalcNeededA2);
            ClassicAssert.IsTrue(recalcNeededA3);
            ClassicAssert.IsTrue(recalcNeededA4);
        }
    }

    [Test]
    public void DeleteWorksheetInvalidatesValues()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet1 = wb.Worksheets.Add("Sheet1");
            IXLWorksheet sheet2 = wb.Worksheets.Add("Sheet2");
            IXLCell sheet1_a1 = sheet1.Cell("A1");
            IXLCell sheet2_a1 = sheet2.Cell("A1");
            sheet1_a1.FormulaA1 = "Sheet2!A1";
            sheet2_a1.Value = "TestValue";

            XLCellValue valueBeforeDeletion = sheet1_a1.Value;
            sheet2.Delete();
            XLCellValue valueAfterDeletion = sheet1_a1.Value;

            ClassicAssert.AreEqual("TestValue", valueBeforeDeletion);
            ClassicAssert.AreEqual(XLError.CellReference, valueAfterDeletion);
        }
    }

    [Test]
    public void CachedValueToExternalWorkbook()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            IXLCell cell = ws.Cell("B2");
            ClassicAssert.IsFalse(cell.NeedsRecalculation);
            ClassicAssert.IsTrue(cell.HasFormula);

            // This will fail when we start supporting external links
            ClassicAssert.IsTrue(cell.FormulaA1.StartsWith("[1]"));

            ClassicAssert.AreEqual("hello world", cell.CachedValue);
            ClassicAssert.AreEqual("hello world", cell.Value);

            ClassicAssert.AreEqual(11, ws.Evaluate("LEN(B2)"));

            NotImplementedException ex = ClassicAssert.Throws<NotImplementedException>(() =>
                wb.RecalculateAllFormulas()
            );
            ClassicAssert.AreEqual(
                "References from other files are not yet implemented.",
                ex.Message
            );
        }
    }

    [Test]
    public void ChangingValueChangesCachedValue()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLCell cell = ws.Cell(1, 1);

            cell.Value = "Hello";
            ClassicAssert.AreEqual("Hello", cell.CachedValue);

            cell.Value = 74.0;
            ClassicAssert.AreEqual(74.0, cell.CachedValue);

            cell.Value = new DateTime(2019, 1, 1, 14, 0, 0);
            ClassicAssert.AreEqual(new DateTime(2019, 1, 1, 14, 0, 0), cell.CachedValue);
        }
    }
}
