using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Hyperlinks;

public class XlHyperlinksTests
{
    [Test]
    [MethodDataSource(nameof(StructuralChangeCases))]
    public void HyperlinkIsMovedOnSheetStructureChange(
        string hyperlinkPosition,
        Action<IXLWorksheet> structuralChange,
        string expectedPosition
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink hyperlink = new("https://example.com");
        ws.Cell(hyperlinkPosition).SetHyperlink(hyperlink);

        structuralChange(ws);

        ClassicAssert.False(ws.Cell(hyperlinkPosition).HasHyperlink);
        ClassicAssert.AreSame(ws.Cell(expectedPosition).GetHyperlink(), hyperlink);
    }

    public static IEnumerable<object[]> StructuralChangeCases =>
        new List<(string, Action<IXLWorksheet>, string)>
        {
            ("D5", ws => ws.Range("A5:B5").Delete(XLShiftDeletedCells.ShiftCellsLeft), "B5"),
            ("D5", ws => ws.Range("B2:D4").Delete(XLShiftDeletedCells.ShiftCellsUp), "D2"),
            ("D5", ws => ws.Column("D").InsertColumnsBefore(2), "F5"), // Insert column leftward
            ("D5", ws => ws.Row(2).InsertRowsAbove(4), "D9"), // Insert row above
        }.Select(x => new object[] { x.Item1, x.Item2, x.Item3 });

    [Test]
    public void ShiftDoesntCollideHyperlinks()
    {
        // In former original data structures, there could be only one hyperlink per area
        // and when links were shifted, one link could shift to a position of another that
        // hasn't been yet shifted. New data structure allows multiple links in a same area,
        // though I hope it's a rare occurence.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        XLHyperlink linkA1 = ws.Cell("A1").CreateHyperlink();
        linkA1.ExternalAddress = new Uri("http://example.com");

        XLHyperlink linkA2 = ws.Cell("A2").CreateHyperlink();
        linkA2.ExternalAddress = new Uri("http://google.com");

        // Original problem was that linkA1 was shifted to A2, but linkA2 wasn't yet shifted to A3.
        // Thus original dictionary threw "An item with the same key has already been added"
        ClassicAssert.DoesNotThrow(() => ws.Row(1).InsertRowsAbove(1));

        ClassicAssert.IsFalse(ws.Cell("A1").HasHyperlink);
        ClassicAssert.AreSame(linkA1, ws.Cell("A2").GetHyperlink());
        ClassicAssert.AreSame(linkA2, ws.Cell("A3").GetHyperlink());
    }

    [Test]
    public void DeleteLinkRemovesLinkFromCell()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        bool deleted = ws.Hyperlinks.Delete(link);

        // Assert
        ClassicAssert.IsTrue(deleted);
        ClassicAssert.AreEqual(ws.Style.Font.FontColor, ws.Cell("A1").Style.Font.FontColor);
        ClassicAssert.AreEqual(ws.Style.Font.Underline, ws.Cell("A1").Style.Font.Underline);
        ClassicAssert.IsNull(link.Container);
        ClassicAssert.IsFalse(ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out _));
    }

    [Test]
    public void DeleteLinkForAddressDeletesTheLink()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        bool deleted = ws.Hyperlinks.Delete(ws.Cell("A1").Address);

        // Assert
        ClassicAssert.IsTrue(deleted);
        ClassicAssert.AreEqual(ws.Style.Font.FontColor, ws.Cell("A1").Style.Font.FontColor);
        ClassicAssert.AreEqual(ws.Style.Font.Underline, ws.Cell("A1").Style.Font.Underline);
        ClassicAssert.IsNull(link.Container);
        ClassicAssert.IsFalse(ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out _));
    }

    [Test]
    public void DeleteLinksForCellAddressWithoutLinkDoesntThrow()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Act
        bool wasDeleted = ws.Hyperlinks.Delete(ws.Cell("A1").Address);

        // Assert
        ClassicAssert.IsFalse(wasDeleted);
    }

    [Test]
    public void DeleteLinkForAddressOfWrongSheetDoesntDeleteTheLink()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        XLHyperlink link = ws1.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        IXLWorksheet ws2 = wb.AddWorksheet();

        // Act
        bool deleted = ws1.Hyperlinks.Delete(ws2.Cell("A1").Address);

        // Assert
        ClassicAssert.IsFalse(deleted);
    }

    [Test]
    public void GetReturnsHyperlinkForAddress()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        XLHyperlink foundLink = ws.Hyperlinks.Get(ws.Cell("A1").Address);

        // Assert
        ClassicAssert.AreSame(link, foundLink);
    }

    [Test]
    public void GetThrowsExceptionWhenAddressDoesntHaveLink()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        IXLWorksheet otherSheet = wb.AddWorksheet();

        // Act + Assert
        foreach (
            IXLAddress addressWithoutLink in new[]
            {
                ws.Cell("A2").Address,
                otherSheet.Cell("A1").Address,
            }
        )
        {
            ClassicAssert.Throws<KeyNotFoundException>(() => ws.Hyperlinks.Get(addressWithoutLink));
        }
    }

    [Test]
    public void GetOnlyReturnsLinksFromCorrectSheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        XLHyperlink link = ws1.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        IXLWorksheet ws2 = wb.AddWorksheet();
        IXLAddress wrongSheetAddress = ws2.Cell("A1").Address;

        ClassicAssert.Throws<KeyNotFoundException>(() => ws1.Hyperlinks.Get(wrongSheetAddress));
    }

    [Test]
    public void TryGetReturnsHyperlinkForAddress()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");

        // Act
        bool wasFound = ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out XLHyperlink? foundLink);

        // Assert
        ClassicAssert.IsTrue(wasFound);
        ClassicAssert.AreSame(link, foundLink);
    }

    [Test]
    public void TryGetDoesntReturnLinkForWrongAddress()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLHyperlink link = ws.Cell("A1").CreateHyperlink();
        link.ExternalAddress = new Uri("https://example.com");
        IXLWorksheet otherSheet = wb.AddWorksheet();

        // Act + Assert
        foreach (
            IXLAddress addressWithoutLink in new[]
            {
                ws.Cell("A2").Address,
                otherSheet.Cell("A1").Address,
            }
        )
        {
            ClassicAssert.IsFalse(ws.Hyperlinks.TryGet(addressWithoutLink, out _));
        }
    }
}
