// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Parser;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.DefinedNames;

[TestFixture]
public class DefinedNamesTests
{
    [Test]
    public void FormulaMustBeValid()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        Assert.Throws<ParsingException>(() => wb.DefinedNames.Add("Test", "SUM(Sheet7!A4"));
    }

    [Test]
    public void CanEvaluateNamedMultiRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
            ws1.Range("A1:C1").Value = 1;
            ws1.Range("A3:C3").Value = 3;
            wb.DefinedNames.Add("TEST", ws1.Ranges("A1:C1,A3:C3"));

            ws1.Cell(2, 1).FormulaA1 = "=SUM(TEST)";

            Assert.AreEqual(12.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
        }
    }

    [Test]
    public void CanGetNamedFromAnother()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").SetValue(1).AddToNamed("value1");

        Assert.AreEqual(1, wb.Cell("value1").Value);
        Assert.AreEqual(1, wb.Range("value1").FirstCell().Value);

        Assert.AreEqual(1, ws1.Cell("value1").Value);
        Assert.AreEqual(1, ws1.Range("value1").FirstCell().Value);

        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");

        ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

        Assert.AreEqual(1, wb.Cell("value2").Value);
        Assert.AreEqual(1, wb.Range("value2").FirstCell().Value);

        Assert.AreEqual(1, ws2.Cell("value1").Value);
        Assert.AreEqual(1, ws2.Range("value1").FirstCell().Value);

        Assert.AreEqual(1, ws2.Cell("value2").Value);
        Assert.AreEqual(1, ws2.Range("value2").FirstCell().Value);
    }

    [Test]
    public void CanGetValidNamedRanges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.Worksheets.Add("Sheet 1");
            IXLWorksheet ws2 = wb.Worksheets.Add("Sheet 2");
            IXLWorksheet ws3 = wb.Worksheets.Add("Sheet'3");

            ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
            ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
            ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
            ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
            wb.DefinedNames.Add(
                "Named range 5",
                new XLRanges(wb) { ws1.Range("A5:D5"), ws3.Range("A5:D5") }
            );

            ws2.Delete();
            ws3.Delete();

            IEnumerable<IXLDefinedName> globalValidRanges = wb.DefinedNames.ValidNamedRanges();
            IEnumerable<IXLDefinedName> globalInvalidRanges = wb.DefinedNames.InvalidNamedRanges();
            IEnumerable<IXLDefinedName> localValidRanges = ws1.DefinedNames.ValidNamedRanges();
            IEnumerable<IXLDefinedName> localInvalidRanges = ws1.DefinedNames.InvalidNamedRanges();

            Assert.AreEqual(1, globalValidRanges.Count());
            Assert.AreEqual("Named range 2", globalValidRanges.First().Name);

            Assert.AreEqual(2, globalInvalidRanges.Count());
            Assert.AreEqual("Named range 4", globalInvalidRanges.First().Name);
            Assert.AreEqual("Named range 5", globalInvalidRanges.Last().Name);

            Assert.AreEqual(1, localValidRanges.Count());
            Assert.AreEqual("Named range 1", localValidRanges.First().Name);

            Assert.AreEqual(0, localInvalidRanges.Count());
        }
    }

    [Test]
    public void CanRenameNamedRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
            IXLDefinedName dn1 = wb.DefinedNames.Add("TEST", "=0.1");

            Assert.IsTrue(wb.DefinedNames.TryGetValue("TEST", out _));
            Assert.IsFalse(wb.DefinedNames.TryGetValue("TEST1", out _));

            dn1.Name = "TEST1";

            Assert.IsFalse(wb.DefinedNames.TryGetValue("TEST", out _));
            Assert.IsTrue(wb.DefinedNames.TryGetValue("TEST1", out _));

            IXLDefinedName dn2 = wb.DefinedNames.Add("TEST2", "=TEST1*2");

            ws1.Cell(1, 1).FormulaA1 = "TEST1";
            ws1.Cell(2, 1).FormulaA1 = "TEST1*10";
            ws1.Cell(3, 1).FormulaA1 = "TEST2";
            ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

            Assert.AreEqual(0.1, (double)ws1.Cell(1, 1).Value, XLHelper.Epsilon);
            Assert.AreEqual(1.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
            Assert.AreEqual(0.2, (double)ws1.Cell(3, 1).Value, XLHelper.Epsilon);
            Assert.AreEqual(0.6, (double)ws1.Cell(4, 1).Value, XLHelper.Epsilon);
        }
    }

    [Test]
    public void CanSaveAndLoadDefinedNames()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet sheet1 = wb.Worksheets.Add("Sheet1");
                IXLWorksheet sheet2 = wb.Worksheets.Add("Sheet2");

                wb.DefinedNames.Add(
                    "wbNamedRange",
                    "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"
                );
                sheet1.DefinedNames.Add(
                    "sheet1NamedRange",
                    "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"
                );
                sheet2.DefinedNames.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet sheet1 = wb.Worksheet("Sheet1");
                IXLWorksheet sheet2 = wb.Worksheet("Sheet2");

                Assert.AreEqual(1, wb.DefinedNames.Count());
                Assert.AreEqual("wbNamedRange", wb.DefinedNames.Single().Name);
                Assert.AreEqual(
                    "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G",
                    wb.DefinedNames.Single().RefersTo
                );
                Assert.AreEqual(5, wb.DefinedNames.Single().Ranges.Count);

                Assert.AreEqual(1, sheet1.DefinedNames.Count());
                Assert.AreEqual("sheet1NamedRange", sheet1.DefinedNames.Single().Name);
                Assert.AreEqual(
                    "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G",
                    sheet1.DefinedNames.Single().RefersTo
                );
                Assert.AreEqual(5, sheet1.DefinedNames.Single().Ranges.Count);

                Assert.AreEqual(1, sheet2.DefinedNames.Count());
                Assert.AreEqual("sheet2NamedRange", sheet2.DefinedNames.Single().Name);
                Assert.AreEqual("Sheet1!A1,Sheet2!A1", sheet2.DefinedNames.Single().RefersTo);
                Assert.AreEqual(2, sheet2.DefinedNames.Single().Ranges.Count);
            }
        }
    }

    [Test]
    public void CopyNamedRangeDifferentWorksheets()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");
        XLRanges ranges = new(wb);
        ranges.Add(ws1.Range("B2:E6"));
        ranges.Add(ws2.Range("D1:E2"));
        IXLDefinedName original = ws1.DefinedNames.Add("Named range", ranges);

        IXLDefinedName copy = original.CopyTo(ws2);

        Assert.AreEqual(1, ws1.DefinedNames.Count());
        Assert.AreEqual(1, ws2.DefinedNames.Count());
        Assert.AreEqual(2, original.Ranges.Count);
        Assert.AreEqual(2, copy.Ranges.Count);
        Assert.AreEqual(original.Name, copy.Name);
        Assert.AreEqual(original.Scope, copy.Scope);
        Assert.AreEqual(
            "Sheet1!B2:E6",
            original.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)
        );
        Assert.AreEqual(
            "Sheet2!D1:E2",
            original.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true)
        );
        Assert.AreEqual(
            "Sheet2!D1:E2",
            copy.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)
        );
        Assert.AreEqual(
            "Sheet2!B2:E6",
            copy.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true)
        );
    }

    [Test]
    public void CopyTableReferencesToDifferentWorksheet()
    {
        // When sheet-scoped name references a table and there is a table with same area in the
        // copied sheet, the copied defined name changes table reference to a new table. If
        // range differs, table reference is not modified.
        using XLWorkbook wb = new();
        IXLWorksheet orgSheet = wb.AddWorksheet();
        orgSheet.Cell("A1").InsertTable(new[] { "Data", "A", "B" }, "OrgTable", true);
        orgSheet.Cell("C1").InsertTable(new[] { "Data", "A", "B" }, "MiscTable", true);
        IXLDefinedName originalName = orgSheet.DefinedNames.Add(
            "TableName",
            "SUM(OrgTable[Data], MiscTable[Data])"
        );

        IXLWorksheet copySheet = wb.AddWorksheet();
        copySheet.Cell("A1").InsertTable(new[] { "Data", "A", "B" }, "CopyTable", true);

        originalName.CopyTo(copySheet);

        IXLDefinedName copyName = copySheet.DefinedNames.Single();
        Assert.AreEqual("TableName", copyName.Name);
        Assert.AreEqual("SUM(CopyTable[Data], MiscTable[Data])", copyName.RefersTo);
    }

    [Test]
    public void CopyWorkbookScopedDefined()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet");
        IXLDefinedName name = wb.DefinedNames.Add("Name", "Sheet!$A$1");

        IXLWorksheet copySheet = wb.AddWorksheet();
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() =>
            name.CopyTo(copySheet)
        )!;
        Assert.AreEqual("Cannot copy workbook scoped defined name.", ex.Message);
    }

    [Test]
    public void CopyDefinedNameToSameSheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Range("B2:E6").AddToNamed("Named range", XLScope.Worksheet);
        IXLDefinedName dn = ws1.DefinedName("Named range");

        TestDelegate action = () => dn.CopyTo(ws1);

        Assert.Throws(typeof(InvalidOperationException), action);
    }

    [Test]
    public void DeleteColumnUsedInNamedRange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Column1");
            ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
            ws.FirstCell().CellRight(2).SetValue("Column3");
            ws.DefinedNames.Add("MyRange", "A1:C1");

            ws.Column(1).Delete();

            Assert.IsTrue(ws.Cell("A1").Style.Font.Bold);
            Assert.AreEqual("Column3", ws.Cell("B1").Value);
            Assert.AreEqual(Blank.Value, ws.Cell("C1").Value);
        }
    }

    [Test]
    public void FormulaIsUpdatedOnSheetRename()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Old name");
        IXLDefinedName bookScopedName = wb.DefinedNames.Add("TEST", "ABS('Old name'!$B$5)");
        IXLDefinedName sheetScopedName = ws.DefinedNames.Add("TEST1", "'Old name'!$D$7:$F$14");

        ws.Name = "Renamed";

        Assert.AreEqual("ABS(Renamed!$B$5)", bookScopedName.RefersTo);
        Assert.AreEqual("Renamed!$B$5:$B$5", bookScopedName.Ranges.ToString());

        Assert.AreEqual("Renamed!$D$7:$F$14", sheetScopedName.RefersTo);
        Assert.AreEqual("Renamed!$D$7:$F$14", sheetScopedName.Ranges.ToString());
    }

    [Test]
    public void MovingRanges()
    {
        XLWorkbook wb = new();

        IXLWorksheet sheet1 = wb.Worksheets.Add("Sheet1");
        IXLWorksheet sheet2 = wb.Worksheets.Add("Sheet2");

        wb.DefinedNames.Add(
            "wbNamedRange",
            "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"
        );
        sheet1.DefinedNames.Add(
            "sheet1NamedRange",
            "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"
        );
        sheet2.DefinedNames.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

        sheet1.Row(1).InsertRowsAbove(2);
        sheet1.Row(1).Delete();
        sheet1.Column(1).InsertColumnsBefore(2);
        sheet1.Column(1).Delete();

        Assert.AreEqual(
            "Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
            wb.DefinedNames.First().RefersTo
        );
        Assert.AreEqual(
            "Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
            sheet1.DefinedNames.First().RefersTo
        );
        Assert.AreEqual("Sheet1!B2,Sheet2!A1", sheet2.DefinedNames.First().RefersTo);

        wb.DefinedNames.ForEach(dn => Assert.AreEqual(XLNamedRangeScope.Workbook, dn.Scope));
        sheet1.DefinedNames.ForEach(dn => Assert.AreEqual(XLNamedRangeScope.Worksheet, dn.Scope));
        sheet2.DefinedNames.ForEach(dn => Assert.AreEqual(XLNamedRangeScope.Worksheet, dn.Scope));
    }

    [Test]
    public void NamedRangeBecomesInvalidOnWorksheetDeleting()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet 1");
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet 2");
        ws1.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
        wb.Ranges("'Sheet 1'!C1:D2,'Sheet 2'!A10:D15").AddToNamed("Compound");

        ws1.Delete();

        Assert.AreEqual(2, wb.DefinedNames.Count());
        Assert.AreEqual(0, wb.DefinedNames.ValidNamedRanges().Count());
        Assert.AreEqual("#REF!", wb.DefinedNames.DefinedName("Simple").RefersTo);
        Assert.AreEqual(
            "#REF!,'Sheet 2'!$A$10:$D$15",
            wb.DefinedNames.DefinedName("Compound").RefersTo
        );
    }

    [Test]
    public void NamedRangeBecomesInvalidOnRangeDeleting()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet 1");
        ws.Range("A1:B2").AddToNamed("Simple");
        ws.Ranges("C1:D2,A10:D15").AddToNamed("Compound");

        ws.Rows(1, 5).Delete();

        Assert.AreEqual(2, wb.DefinedNames.Count());
        Assert.AreEqual(0, wb.DefinedNames.ValidNamedRanges().Count());
        Assert.AreEqual("#REF!", wb.DefinedNames.DefinedName("Simple").RefersTo);
        Assert.AreEqual(
            "#REF!,'Sheet 1'!$A$5:$D$10",
            wb.DefinedNames.DefinedName("Compound").RefersTo
        );
    }

    [Test]
    public void NamedRangeMayReferToExpression()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
                wb.DefinedNames.Add("TEST", "=0.1");
                wb.DefinedNames.Add("TEST2", "=TEST*2");

                ws1.Cell(1, 1).FormulaA1 = "TEST";
                ws1.Cell(2, 1).FormulaA1 = "TEST*10";
                ws1.Cell(3, 1).FormulaA1 = "TEST2";
                ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

                Assert.AreEqual(0.1, (double)ws1.Cell(1, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(1.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(0.2, (double)ws1.Cell(3, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(0.6, (double)ws1.Cell(4, 1).Value, XLHelper.Epsilon);

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws1 = wb.Worksheets.First();

                Assert.AreEqual(0.1, (double)ws1.Cell(1, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(1.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(0.2, (double)ws1.Cell(3, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(0.6, (double)ws1.Cell(4, 1).Value, XLHelper.Epsilon);
            }
        }
    }

    [Test]
    public void NamedRangeReferringToMultipleRangesCanBeSavedAndLoaded()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet 1");

                wb.DefinedNames.Add(
                    "Multirange named range",
                    new XLRanges(wb) { ws.Range("A5:D5"), ws.Range("A15:D15") }
                );

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                Assert.AreEqual(1, wb.DefinedNames.Count());
                XLDefinedName nr = (XLDefinedName)wb.DefinedNames.Single();
                Assert.AreEqual("'Sheet 1'!$A$5:$D$5,'Sheet 1'!$A$15:$D$15", nr.RefersTo);
                Assert.AreEqual(2, nr.Ranges.Count);
                Assert.AreEqual(
                    "'Sheet 1'!A5:D5",
                    nr.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)
                );
                Assert.AreEqual(
                    "'Sheet 1'!A15:D15",
                    nr.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true)
                );
                Assert.AreEqual(2, nr.SheetReferencesList.Count);
                Assert.AreEqual("'Sheet 1'!$A$5:$D$5", nr.SheetReferencesList.First());
                Assert.AreEqual("'Sheet 1'!$A$15:$D$15", nr.SheetReferencesList.Last());
            }
        }
    }

    [Test]
    public void DefinedNamesReferencingSheetRangeBecomeInvalidWhenSheetIsDeleted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.Worksheets.Add("Sheet 1");
            IXLWorksheet ws2 = wb.Worksheets.Add("Sheet 2");
            IXLWorksheet ws3 = wb.Worksheets.Add("Sheet'3");

            ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
            ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
            ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
            ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
            wb.DefinedNames.Add(
                "Named range 5",
                new XLRanges(wb) { ws1.Range("A5:D5"), ws3.Range("A5:D5") }
            );

            ws2.Delete();
            ws3.Delete();

            Assert.AreEqual(1, ws1.DefinedNames.Count());
            Assert.AreEqual("Named range 1", ws1.DefinedNames.First().Name);
            Assert.AreEqual(XLNamedRangeScope.Worksheet, ws1.DefinedNames.First().Scope);
            Assert.AreEqual("'Sheet 1'!$A$1:$D$1", ws1.DefinedNames.First().RefersTo);
            Assert.AreEqual(
                "'Sheet 1'!A1:D1",
                ws1.DefinedNames.First()
                    .Ranges.Single()
                    .RangeAddress.ToString(XLReferenceStyle.A1, true)
            );

            Assert.AreEqual(3, wb.DefinedNames.Count());

            Assert.AreEqual("Named range 2", wb.DefinedNames.ElementAt(0).Name);
            Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(0).Scope);
            Assert.AreEqual("'Sheet 1'!$A$2:$D$2", wb.DefinedNames.ElementAt(0).RefersTo);
            Assert.AreEqual(
                "'Sheet 1'!A2:D2",
                wb.DefinedNames.ElementAt(0)
                    .Ranges.Single()
                    .RangeAddress.ToString(XLReferenceStyle.A1, true)
            );

            Assert.AreEqual("Named range 4", wb.DefinedNames.ElementAt(1).Name);
            Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(1).Scope);
            Assert.AreEqual("#REF!", wb.DefinedNames.ElementAt(1).RefersTo);
            Assert.IsFalse(wb.DefinedNames.ElementAt(1).Ranges.Any());

            Assert.AreEqual("Named range 5", wb.DefinedNames.ElementAt(2).Name);
            Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(2).Scope);
            Assert.AreEqual("'Sheet 1'!$A$5:$D$5,#REF!", wb.DefinedNames.ElementAt(2).RefersTo);
            Assert.AreEqual(1, wb.DefinedNames.ElementAt(2).Ranges.Count);
            Assert.AreEqual(
                "'Sheet 1'!A5:D5",
                wb.DefinedNames.ElementAt(2)
                    .Ranges.Single()
                    .RangeAddress.ToString(XLReferenceStyle.A1, true)
            );
        }
    }

    [Test]
    public void NamedRangesFromDeletedSheetAreSavedWithoutAddress()
    {
        // Range address referring to the deleted sheet look like #REF!A1:B2.
        // But workbooks with such references in named ranges Excel considers as broken files.
        // It requires #REF!

        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                wb.Worksheets.Add("Sheet 1");
                IXLWorksheet ws2 = wb.Worksheets.Add("Sheet 2");
                ws2.Range("A4:D4").AddToNamed("Test named range", XLScope.Workbook);
                ws2.Delete();
                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                Assert.AreEqual("#REF!", wb.DefinedNames.Single().RefersTo);
            }
        }
    }

    [Test]
    public void OnlyWorksheetScopedDefinedNamesAreCopiedWhenSheetIsCopied()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
            ws1.FirstCell().InsertData(Enumerable.Range(1, 10));
            wb.DefinedNames.Add("wbNamedRange", ws1.Range("A1:A10"));
            ws1.DefinedNames.Add("wsNamedRange", ws1.Range("A3"));

            IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");
            ws2.FirstCell().InsertData(Enumerable.Range(101, 10));
            ws1.DefinedNames.Add("wsNamedRangeAcrossSheets", ws2.Range("A4"));

            ws1.Cell("C1").FormulaA1 = "=wbNamedRange";
            ws1.Cell("C2").FormulaA1 = "=wsNamedRange";
            ws1.Cell("C3").FormulaA1 = "=wsNamedRangeAcrossSheets";

            Assert.AreEqual(1, ws1.Cell("C1").Value);
            Assert.AreEqual(3, ws1.Cell("C2").Value);
            Assert.AreEqual(104, ws1.Cell("C3").Value);

            IXLWorksheet wsCopy = ws1.CopyTo("Copy");
            Assert.AreEqual(1, wsCopy.Cell("C1").Value);
            Assert.AreEqual(3, wsCopy.Cell("C2").Value);
            Assert.AreEqual(104, wsCopy.Cell("C3").Value);

            Assert.AreEqual(
                "Sheet1!A1:A10",
                wb.DefinedName("wbNamedRange").Ranges.First().RangeAddress.ToStringRelative(true)
            );
            Assert.AreEqual(
                "Copy!A3:A3",
                wsCopy
                    .DefinedName("wsNamedRange")
                    .Ranges.First()
                    .RangeAddress.ToStringRelative(true)
            );
            Assert.AreEqual(
                "Sheet2!A4:A4",
                wsCopy
                    .DefinedName("wsNamedRangeAcrossSheets")
                    .Ranges.First()
                    .RangeAddress.ToStringRelative(true)
            );
        }
    }

    [Test]
    public void SavedDefinedNamesBecomeInvalidOnSheetDeleting()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws1 = wb.Worksheets.Add("Sheet 1");
                IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");
                IXLWorksheet ws3 = wb.Worksheets.Add("Sheet'3");

                ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
                ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
                ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
                ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
                wb.DefinedNames.Add(
                    "Named range 5",
                    new XLRanges(wb) { ws1.Range("A5:D5"), ws3.Range("A5:D5") }
                );

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                wb.Worksheet("Sheet2").Delete();
                wb.Worksheet("Sheet'3").Delete();
                wb.Save();
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws1 = wb.Worksheet("Sheet 1");
                Assert.AreEqual(1, ws1.DefinedNames.Count());
                Assert.AreEqual("Named range 1", ws1.DefinedNames.First().Name);
                Assert.AreEqual(XLNamedRangeScope.Worksheet, ws1.DefinedNames.First().Scope);
                Assert.AreEqual("'Sheet 1'!$A$1:$D$1", ws1.DefinedNames.First().RefersTo);
                Assert.AreEqual(
                    "'Sheet 1'!A1:D1",
                    ws1.DefinedNames.First()
                        .Ranges.Single()
                        .RangeAddress.ToString(XLReferenceStyle.A1, true)
                );

                Assert.AreEqual(3, wb.DefinedNames.Count());

                Assert.AreEqual("Named range 2", wb.DefinedNames.ElementAt(0).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(0).Scope);
                Assert.AreEqual("'Sheet 1'!$A$2:$D$2", wb.DefinedNames.ElementAt(0).RefersTo);
                Assert.AreEqual(
                    "'Sheet 1'!A2:D2",
                    wb.DefinedNames.ElementAt(0)
                        .Ranges.Single()
                        .RangeAddress.ToString(XLReferenceStyle.A1, true)
                );

                Assert.AreEqual("Named range 4", wb.DefinedNames.ElementAt(1).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(1).Scope);
                Assert.AreEqual("#REF!", wb.DefinedNames.ElementAt(1).RefersTo);
                Assert.IsFalse(wb.DefinedNames.ElementAt(1).Ranges.Any());

                Assert.AreEqual("Named range 5", wb.DefinedNames.ElementAt(2).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(2).Scope);
                Assert.AreEqual("'Sheet 1'!$A$5:$D$5,#REF!", wb.DefinedNames.ElementAt(2).RefersTo);
                Assert.AreEqual(1, wb.DefinedNames.ElementAt(2).Ranges.Count);
                Assert.AreEqual(
                    "'Sheet 1'!A5:D5",
                    wb.DefinedNames.ElementAt(2)
                        .Ranges.Single()
                        .RangeAddress.ToString(XLReferenceStyle.A1, true)
                );
            }
        }
    }

    [Test]
    public void TestInvalidNamedRangeOnWorkbookScope()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Column1");
            ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
            ws.FirstCell().CellRight(2).SetValue("Column3");

            Assert.Throws<ArgumentException>(() => wb.DefinedNames.Add("MyRange", "A1:C1"));
        }
    }

    [Test]
    public void WbContainsWsNamedRange()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

        Assert.IsTrue(wb.DefinedNames.Contains("Sheet1!Name"));
        Assert.IsFalse(wb.DefinedNames.Contains("Sheet1!NameX"));

        Assert.IsNotNull(wb.DefinedName("Sheet1!Name"));
        Assert.IsNull(wb.DefinedName("Sheet1!NameX"));

        Boolean found1 = wb.DefinedNames.TryGetValue(
            "Sheet1!Name",
            out IXLDefinedName? definedName1
        );
        Assert.IsTrue(found1);
        Assert.IsNotNull(definedName1);
        Assert.AreEqual(XLNamedRangeScope.Worksheet, definedName1.Scope);

        Boolean found2 = wb.DefinedNames.TryGetValue(
            "Sheet1!NameX",
            out IXLDefinedName? definedName2
        );
        Assert.IsFalse(found2);
        Assert.IsNull(definedName2);
    }

    [Test]
    public void WorkbookContainsNamedRange()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddToNamed("Name");

        Assert.IsTrue(wb.DefinedNames.Contains("Name"));
        Assert.IsFalse(wb.DefinedNames.Contains("NameX"));

        Assert.IsNotNull(wb.DefinedName("Name"));
        Assert.IsNull(wb.DefinedName("NameX"));

        Boolean found1 = wb.DefinedNames.TryGetValue("Name", out IXLDefinedName? definedName1);
        Assert.IsTrue(found1);
        Assert.IsNotNull(definedName1);

        Boolean found2 = wb.DefinedNames.TryGetValue("NameX", out IXLDefinedName? definedName2);
        Assert.IsFalse(found2);
        Assert.IsNull(definedName2);
    }

    [Test]
    public void WorksheetContainsNamedRange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

        Assert.IsTrue(ws.DefinedNames.Contains("Name"));
        Assert.IsFalse(ws.DefinedNames.Contains("NameX"));

        Assert.IsNotNull(ws.DefinedName("Name"));
        Assert.Throws<KeyNotFoundException>(() => ws.DefinedName("NameX"));

        Boolean found1 = ws.DefinedNames.TryGetValue("Name", out IXLDefinedName? definedName1);
        Assert.IsTrue(found1);
        Assert.IsNotNull(definedName1);

        Boolean found2 = ws.DefinedNames.TryGetValue("NameX", out IXLDefinedName? definedName2);
        Assert.IsFalse(found2);
        Assert.IsNull(definedName2);
    }

    [Test]
    public void NamedRangeWithSameNameAsAFunction()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLCell a1 = ws.FirstCell();
        IXLCell a2 = a1.CellBelow();

        a1.SetValue(5).AddToNamed("RAND");
        a2.FormulaA1 = "=RAND * 10";

        Assert.AreEqual(50, a2.GetDouble());
    }

    [Test]
    public void RefersToThrowsOnNull()
    {
        using XLWorkbook wb = new();
        IXLDefinedName name = wb.DefinedNames.Add("name", "1+2");
        Assert.Throws<ArgumentNullException>(() => name.RefersTo = null!);
    }

    [TestCase("")]
    [TestCase("=  ")]
    [TestCase("  ")]
    public void RefersToCantBeEmpty(string formula)
    {
        // Excel will try to repair a workbook that contains a defined name with a formula that is an empty string.
        using XLWorkbook wb = new();
        IXLDefinedName name = wb.DefinedNames.Add("demo", "1+2");
        const string message = "Formula can't be empty.";

        Assert.That(
            () => name.SetRefersTo(formula),
            Throws.Exception.TypeOf<ArgumentException>().With.Message.EqualTo(message)
        );
        Assert.That(
            () => wb.DefinedNames.Add("name", formula),
            Throws.Exception.TypeOf<ArgumentException>().With.Message.EqualTo(message)
        );
    }
}
