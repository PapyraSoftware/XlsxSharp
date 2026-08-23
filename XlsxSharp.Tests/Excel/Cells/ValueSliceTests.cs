using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Cells;

public class ValueSliceTests
{
    [Test]
    public void DeletingWorksheetDereferencesAllTextsInItsValueSlice()
    {
        using XLWorkbook wb = new();
        SharedStringTable sst = wb.SharedStringTable;
        IXLWorksheet keptWs = wb.AddWorksheet();
        IXLWorksheet removedWs = wb.AddWorksheet();
        keptWs.Cell("A1").Value = "Double referenced text";
        removedWs.Cell("A1").Value = "Double referenced text";
        removedWs.Cell("B1").Value = "Single referenced text";

        ClassicAssert.AreEqual(2, sst.Count);

        wb.Worksheets.Delete(removedWs.Name);

        ClassicAssert.AreEqual(1, sst.Count);
        ClassicAssert.AreEqual("Double referenced text", keptWs.Cell(1, 1).Value);
    }

    [Test]
    public void ClearDereferencesTextsInTheRange()
    {
        using XLWorkbook wb = new();
        SharedStringTable sst = wb.SharedStringTable;
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Double referenced text";
        ws.Cell("B2").Value = "Double referenced text";
        ws.Cell("C2").Value = "Single referenced text";

        ClassicAssert.AreEqual(2, sst.Count);
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.Clear(new Area(2, 2, 2, 3));
        ClassicAssert.AreEqual(1, sst.Count);
        ClassicAssert.AreEqual("Double referenced text", ws.Cell("A1").Value);
    }

    [Test]
    public void DeleteAreaAndShiftLeftDereferencesAllTextsDeletedArea()
    {
        using XLWorkbook wb = new();
        SharedStringTable sst = wb.SharedStringTable;
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Deleted Single Reference"; // id 0
        ws.Cell("C1").Value = "Kept Single Reference"; // id 1
        ws.Cell("B1").Value = "Kept Double Reference"; // id 2
        ws.Cell("C3").Value = "Kept Double Reference"; // id 2

        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.DeleteAreaAndShiftLeft(
            new Area(2, 2, 3, 3)
        );

        ClassicAssert.AreEqual(2, sst.Count);
        ClassicAssert.AreEqual("Kept Single Reference", sst[1]);
        ClassicAssert.AreEqual("Kept Double Reference", sst[2]);
    }

    [Test]
    public void DeleteAreaAndShiftUpDereferencesAllTextsDeletedArea()
    {
        using XLWorkbook wb = new();
        SharedStringTable sst = wb.SharedStringTable;
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Deleted Single Reference"; // id 0
        ws.Cell("A3").Value = "Kept Single Reference"; // id 1
        ws.Cell("A2").Value = "Kept Double Reference"; // id 2
        ws.Cell("C3").Value = "Kept Double Reference"; // id 2

        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.DeleteAreaAndShiftLeft(
            new Area(2, 2, 3, 3)
        );

        ClassicAssert.AreEqual(2, sst.Count);
        ClassicAssert.AreEqual("Kept Single Reference", sst[1]);
        ClassicAssert.AreEqual("Kept Double Reference", sst[2]);
    }

    [Test]
    public void InsertAreaAndShiftDownDereferencesAllTextsInPushedOutRange()
    {
        using XLWorkbook wb = new();
        SharedStringTable sst = wb.SharedStringTable;
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("B2").Value = "Kept Single Reference"; // id 0
        ws.Cell("C1048576").Value = "Deleted Single Reference"; // id 1
        ws.Cell("B1048574").Value = "Kept Double Reference"; // id 2
        ws.Cell("B1048576").Value = "Kept Double Reference"; // id 2
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.InsertAreaAndShiftDown(
            new Area(3, 2, 4, 3)
        );

        ClassicAssert.AreEqual(2, sst.Count);
        ClassicAssert.AreEqual("Kept Single Reference", sst[0]);
        ClassicAssert.AreEqual("Kept Double Reference", sst[2]);
    }

    [Test]
    public void InsertAreaAndShiftRightDereferencesAllTextsInPushedOutRange()
    {
        using XLWorkbook wb = new();
        SharedStringTable sst = wb.SharedStringTable;
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("B2").Value = "Kept Single Reference"; // id 0
        ws.Cell("XFD2").Value = "Deleted Single Reference"; // id 1
        ws.Cell("XFD3").Value = "Kept Double Reference"; // id 2
        ws.Cell("XFB3").Value = "Kept Double Reference"; // id 2
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.InsertAreaAndShiftRight(
            new Area(2, 3, 3, 4)
        );

        ClassicAssert.AreEqual(2, sst.Count);
        ClassicAssert.AreEqual("Kept Single Reference", sst[0]);
        ClassicAssert.AreEqual("Kept Double Reference", sst[2]);
    }
}
