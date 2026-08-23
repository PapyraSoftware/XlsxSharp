using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Tests.Excel.Ranges;

public class SortTests
{
    [Test]
    public void ValuesAreSortedByTypeFirst()
    {
        // The values in asc order are number, text, logical, error, blanks.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue[] values =
        [
            1,
            "",
            "#VALUE!",
            "1",
            "Text",
            "TRUE",
            true,
            XLError.IncompatibleValue,
            Blank.Value,
        ];

        // Assign in reverse order
        for (int row = 1; row <= values.Length; ++row)
        {
            ws.Cell(row, 1).Value = values[^row];
        }

        ws.Range(1, 1, values.Length, 1).Sort("1 ASC");

        for (int row = 1; row <= values.Length; ++row)
        {
            XLCellValue sortedValue = ws.Cell(row, 1).Value;
            ClassicAssert.AreEqual(values[row - 1], sortedValue);
        }
    }

    [Test]
    [Arguments(XLSortOrder.Ascending)]
    [Arguments(XLSortOrder.Descending)]
    public void BlanksAreAlwaysLast(XLSortOrder sortOrder)
    {
        // When range contains blank, it is always last, no matter
        // if the sort order is ascending or descending
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue[] values = [1, Blank.Value, 2];
        for (int row = 1; row <= values.Length; ++row)
        {
            ws.Cell(row, 1).Value = values[row - 1];
        }

        ws.Range(1, 1, values.Length, 1).Sort("1", sortOrder);

        ClassicAssert.AreEqual(Blank.Value, ws.Cell(3, 1).Value);
    }

    [Test]
    public void IgnoreBlanksSetToFalseTreatsBlanksAsEmptyStrings()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "Text";
        ws.Cell("A2").Value = Blank.Value;
        ws.Cell("A3").Value = string.Empty;

        ws.Range("A1:A3").Sort(1, ignoreBlanks: false);

        // Since blank is treated as empty string, it is not shuffled to the end.
        ClassicAssert.AreEqual(Blank.Value, ws.Cell("A1").Value);
        ClassicAssert.AreEqual(string.Empty, ws.Cell("A2").Value);
        ClassicAssert.AreEqual("Text", ws.Cell("A3").Value);
    }

    [Test]
    [Arguments(true, "a", "A")]
    [Arguments(false, "A", "a")]
    public void MatchCaseFlagDeterminesIfTextsAreComparedCaseSensitive(
        bool matchCase,
        string expectedFirst,
        string expectedSecond
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // In US locale, lower-case is before upper case.
        ws.Cell("A1").Value = "A";
        ws.Cell("A2").Value = "a";

        ws.Range("A1:A2").Sort(1, matchCase: matchCase);

        ClassicAssert.AreEqual(expectedFirst, ws.Cell("A1").Value);
        ClassicAssert.AreEqual(expectedSecond, ws.Cell("A2").Value);
    }

    [Test]
    public void SortCanUseMultipleColumns()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell().InsertData(new object[] { new[] { 1, 2 }, new[] { 2, 2 }, new[] { 1, 1 } });

        ws.Range("A1:B4").Sort("2 ASC, 1 DESC");

        ClassicAssert.AreEqual(1, ws.Cell("A1").Value);
        ClassicAssert.AreEqual(1, ws.Cell("B1").Value);
        ClassicAssert.AreEqual(2, ws.Cell("A2").Value);
        ClassicAssert.AreEqual(2, ws.Cell("B2").Value);
        ClassicAssert.AreEqual(1, ws.Cell("A3").Value);
        ClassicAssert.AreEqual(2, ws.Cell("B3").Value);
    }

    [Test]
    public void SortColumnsInRangeByRows()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell().InsertData(new object[] { new[] { 2, 2, 1 }, new[] { 1, 2, 1 } });

        // Doesn't have parameters, so it is first rows ASC, second row ASC.
        ws.Range("A1:C2").SortLeftToRight();

        ClassicAssert.AreEqual(1, ws.Cell("A1").Value);
        ClassicAssert.AreEqual(1, ws.Cell("A2").Value);
        ClassicAssert.AreEqual(2, ws.Cell("B1").Value);
        ClassicAssert.AreEqual(1, ws.Cell("B2").Value);
        ClassicAssert.AreEqual(2, ws.Cell("C1").Value);
        ClassicAssert.AreEqual(2, ws.Cell("C2").Value);
    }
}
