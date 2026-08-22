using System;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Protection;

namespace XlsxSharp.Tests.Excel.Misc;

[TestFixture]
public class XLWorkbookTests
{
    [Test]
    public void Cell1()
    {
        XLWorkbook wb = new();
        IXLCell cell = wb.Cell("ABC");
        Assert.IsNull(cell);
    }

    [Test]
    public void Cell2()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        IXLCell cell = wb.Cell("Sheet1!Result");
        Assert.IsNotNull(cell);
        Assert.AreEqual(1, cell.Value);
    }

    [Test]
    public void Cell3()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        IXLCell cell = wb.Cell("Sheet1!Result");
        Assert.IsNotNull(cell);
        Assert.AreEqual(1, cell.Value);
    }

    [Test]
    public void Cells1()
    {
        XLWorkbook wb = new();
        IXLCells cells = wb.Cells("ABC");
        Assert.IsNotNull(cells);
        Assert.AreEqual(0, cells.Count());
    }

    [Test]
    public void Cells2()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        IXLCells cells = wb.Cells("Sheet1!Result, ABC");
        Assert.IsNotNull(cells);
        Assert.AreEqual(1, cells.Count());
        Assert.AreEqual(1, cells.First().Value);
    }

    [Test]
    public void Cells3()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        IXLCells cells = wb.Cells("Sheet1!Result, ABC");
        Assert.IsNotNull(cells);
        Assert.AreEqual(1, cells.Count());
        Assert.AreEqual(1, cells.First().Value);
    }

    [Test]
    public void GetCellFromFullAddress()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLWorksheet ws2 = wb.AddWorksheet("O'Sheet 2");
        IXLCell c1 = ws.Cell("C123");
        IXLCell c2 = ws2.Cell("B7");

        IXLCell c1_full = wb.Cell("Sheet1!C123");
        IXLCell c2_full = wb.Cell("'O'Sheet 2'!B7");

        Assert.AreEqual(c1, c1_full);
        Assert.AreEqual(c2, c2_full);
        Assert.NotNull(c1_full);
        Assert.NotNull(c2_full);
    }

    [TestCase("Sheet1")]
    [TestCase("Sheet1!")]
    [TestCase("Sheet2!")]
    [TestCase("Sheet2!C1")]
    [TestCase("Sheet1!ZZZ1")]
    [TestCase("Sheet1!A")]
    public void GetCellFromNonExistingFullAddress(string address)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");

        IXLCell c = wb.Cell(address);

        Assert.IsNull(c);
    }

    [Test]
    public void GetRangeFromFullAddress()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLRange r1 = ws.Range("C123:D125");

        IXLRange r2 = wb.Range("Sheet1!C123:D125");

        Assert.AreSame(r1, r2);
        Assert.NotNull(r2);
    }

    [TestCase("Sheet2!C1:D2")]
    [TestCase("Sheet1!A")]
    public void GetRangeFromNonExistingFullAddress(string rangeAddress)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");

        IXLRange r = wb.Range(rangeAddress);

        Assert.IsNull(r);
    }

    [Test]
    public void GetRangesFromFullAddress()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLRanges r1 = ws.Ranges("A1:B2,C1:E3");

        IXLRanges r2 = wb.Ranges("Sheet1!A1:B2,Sheet1!C1:E3");

        Assert.AreEqual(2, r2.Count);
        Assert.AreSame(r1.First(), r2.First());
        Assert.AreSame(r1.Last(), r2.Last());
    }

    [TestCase("Sheet2!C1:D2,Sheet2!F1:G4")]
    [TestCase("Sheet1!A,Sheet1!B")]
    public void GetRangesFromNonExistingFullAddress(string rangesAddress)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");

        IXLRanges r = wb.Ranges(rangesAddress);

        Assert.NotNull(r);
        Assert.False(r.Any());
    }

    [Test]
    public void NonExistentDefinedNameReturnsNull()
    {
        XLWorkbook wb = new();
        IXLDefinedName? definedName = wb.DefinedName("ABC");
        Assert.IsNull(definedName);
    }

    [Test]
    public void SheetSpecifiedDefinedNameIsRetrievedFromSheetIfDefinedThere()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        IXLDefinedName? definedName = wb.DefinedName("Sheet1!Result");
        Assert.IsNotNull(definedName);
        Assert.AreEqual(1, definedName.Ranges.Count);
        Assert.AreEqual(1, definedName.Ranges.Cells().Count());
        Assert.AreEqual(1, definedName.Ranges.First().FirstCell().Value);
    }

    [Test]
    public void SheetSpecifiedDefinedNameReturnsNullIfNotDefinedInSheetNorWorkbook()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLDefinedName? definedName = wb.DefinedName("Sheet1!Result");
        Assert.IsNull(definedName);
    }

    [Test]
    public void SheetSpecifiedDefinedNameFallsBackToWorkbookScopedDefinedNameIfNotDefinedInSheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        IXLDefinedName? definedName = wb.DefinedName("Sheet1!Result");
        Assert.IsNotNull(definedName);
        Assert.AreEqual(1, definedName.Ranges.Count);
        Assert.AreEqual(1, definedName.Ranges.Cells().Count());
        Assert.AreEqual(1, definedName.Ranges.First().FirstCell().Value);
    }

    [Test]
    public void Range1()
    {
        XLWorkbook wb = new();
        IXLRange range = wb.Range("ABC");
        Assert.IsNull(range);
    }

    [Test]
    public void Range2()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        IXLRange range = wb.Range("Sheet1!Result");
        Assert.IsNotNull(range);
        Assert.AreEqual(1, range.Cells().Count());
        Assert.AreEqual(1, range.FirstCell().Value);
    }

    [Test]
    public void Range3()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        IXLRange range = wb.Range("Sheet1!Result");
        Assert.IsNotNull(range);
        Assert.AreEqual(1, range.Cells().Count());
        Assert.AreEqual(1, range.FirstCell().Value);
    }

    [Test]
    public void Ranges1()
    {
        XLWorkbook wb = new();
        IXLRanges ranges = wb.Ranges("ABC");
        Assert.IsNotNull(ranges);
        Assert.AreEqual(0, ranges.Count);
    }

    [Test]
    public void Ranges2()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        IXLRanges ranges = wb.Ranges("Sheet1!Result, ABC");
        Assert.IsNotNull(ranges);
        Assert.AreEqual(1, ranges.Cells().Count());
        Assert.AreEqual(1, ranges.First().FirstCell().Value);
    }

    [Test]
    public void Ranges3()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        IXLRanges ranges = wb.Ranges("Sheet1!Result, ABC");
        Assert.IsNotNull(ranges);
        Assert.AreEqual(1, ranges.Cells().Count());
        Assert.AreEqual(1, ranges.First().FirstCell().Value);
    }

    [Test]
    public void WbNamedCell()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
        Assert.AreEqual("Test", wb.Cell("TestCell").GetText());
        Assert.AreEqual("Test", ws.Cell("TestCell").GetText());
    }

    [Test]
    public void WbNamedCells()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
        ws.Cell(2, 1).SetValue("B").AddToNamed("Test2");
        IXLCells wbCells = wb.Cells("TestCell, Test2");
        Assert.AreEqual("Test", wbCells.First().GetText());
        Assert.AreEqual("B", wbCells.Last().GetText());

        IXLCells wsCells = ws.Cells("TestCell, Test2");
        Assert.AreEqual("Test", wsCells.First().GetText());
        Assert.AreEqual("B", wsCells.Last().GetText());
    }

    [Test]
    public void WbNamedRange()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("A");
        ws.Cell(2, 1).SetValue("B");
        IXLRange original = ws.Range("A1:A2");
        original.AddToNamed("TestRange");
        Assert.AreEqual(
            original.RangeAddress.ToStringFixed(),
            wb.Range("TestRange").RangeAddress.ToString()
        );
        Assert.AreEqual(
            original.RangeAddress.ToStringFixed(),
            ws.Range("TestRange").RangeAddress.ToString()
        );
    }

    [Test]
    public void WbNamedRanges()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("A");
        ws.Cell(2, 1).SetValue("B");
        ws.Cell(3, 1).SetValue("C").AddToNamed("Test2");
        IXLRange original = ws.Range("A1:A2");
        original.AddToNamed("TestRange");
        IXLRanges wbRanges = wb.Ranges("TestRange, Test2");
        Assert.AreEqual(
            original.RangeAddress.ToStringFixed(),
            wbRanges.First().RangeAddress.ToString()
        );
        Assert.AreEqual("$A$3:$A$3", wbRanges.Last().RangeAddress.ToStringFixed());

        IXLRanges wsRanges = wb.Ranges("TestRange, Test2");
        Assert.AreEqual(
            original.RangeAddress.ToStringFixed(),
            wsRanges.First().RangeAddress.ToString()
        );
        Assert.AreEqual("$A$3:$A$3", wsRanges.Last().RangeAddress.ToStringFixed());
    }

    [Test]
    public void WbNamedRangesOneString()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        wb.DefinedNames.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");

        IXLRanges wbRanges = ws.Ranges("TestRange");
        Assert.AreEqual("$A$1:$A$1", wbRanges.First().RangeAddress.ToStringFixed());
        Assert.AreEqual("$A$3:$A$3", wbRanges.Last().RangeAddress.ToStringFixed());

        IXLRanges wsRanges = ws.Ranges("TestRange");
        Assert.AreEqual("$A$1:$A$1", wsRanges.First().RangeAddress.ToStringFixed());
        Assert.AreEqual("$A$3:$A$3", wsRanges.Last().RangeAddress.ToStringFixed());
    }

    [Test]
    public void WbProtect1()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.Protect();
            Assert.IsTrue(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsFalse(wb.IsPasswordProtected);
        }
    }

    [Test]
    public void WbProtect2()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.Protect(XLWorkbookProtectionElements.Windows);
            Assert.IsTrue(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsFalse(wb.IsPasswordProtected);
        }
    }

    [Test]
    public void WbProtect3()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.Protect("Abc@123");
            Assert.IsTrue(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsTrue(wb.IsPasswordProtected);
            Assert.Throws<InvalidOperationException>(() => wb.Protect());
            Assert.Throws<InvalidOperationException>(() => wb.Unprotect());
            Assert.Throws<ArgumentException>(() => wb.Unprotect("Cde@345"));
        }
    }

    [Test]
    public void WbProtect4()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.Protect();
            Assert.IsTrue(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsFalse(wb.IsPasswordProtected);
            wb.Unprotect();
            wb.Protect("Abc@123");
            Assert.IsTrue(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsTrue(wb.IsPasswordProtected);
        }
    }

    [Test]
    public void WbProtect5()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.Protect(
                "Abc@123",
                XLProtectionAlgorithm.DefaultProtectionAlgorithm,
                XLWorkbookProtectionElements.Windows
            );
            Assert.IsTrue(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsTrue(wb.IsPasswordProtected);
            wb.Unprotect("Abc@123");
            Assert.IsFalse(wb.LockStructure);
            Assert.IsFalse(wb.LockWindows);
            Assert.IsFalse(wb.IsPasswordProtected);
        }
    }

    [Test]
    public void FileSharingProperties()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                wb.AddWorksheet("Sheet1").Cell("A1").Value = "Hello world!";
                wb.FileSharing.ReadOnlyRecommended = true;
                wb.FileSharing.UserName = Environment.UserName;
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                Assert.IsTrue(wb.FileSharing.ReadOnlyRecommended);
                Assert.AreEqual(Environment.UserName, wb.FileSharing.UserName);
            }
        }
    }

    [Test]
    public void AccessDisposedWorkbookThrowsException()
    {
        IXLWorkbook wb;
        using (wb = new XLWorkbook())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world");
        }

        Assert.Throws<ObjectDisposedException>(() =>
            Console.WriteLine(wb.Worksheets.First().FirstCell().Value)
        );
    }
}
