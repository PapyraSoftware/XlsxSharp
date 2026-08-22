using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Hyperlinks;

[TestFixture]
[TestOf(typeof(XLHyperlinks))]
public class XLHyperlinksTests
{
    [TestCaseSource(nameof(StructuralChangeCases))]
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

        Assert.False(ws.Cell(hyperlinkPosition).HasHyperlink);
        Assert.AreSame(ws.Cell(expectedPosition).GetHyperlink(), hyperlink);
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
        Assert.DoesNotThrow(() => ws.Row(1).InsertRowsAbove(1));

        Assert.IsFalse(ws.Cell("A1").HasHyperlink);
        Assert.AreSame(linkA1, ws.Cell("A2").GetHyperlink());
        Assert.AreSame(linkA2, ws.Cell("A3").GetHyperlink());
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
        Assert.IsTrue(deleted);
        Assert.AreEqual(ws.Style.Font.FontColor, ws.Cell("A1").Style.Font.FontColor);
        Assert.AreEqual(ws.Style.Font.Underline, ws.Cell("A1").Style.Font.Underline);
        Assert.IsNull(link.Container);
        Assert.IsFalse(ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out _));
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
        Assert.IsTrue(deleted);
        Assert.AreEqual(ws.Style.Font.FontColor, ws.Cell("A1").Style.Font.FontColor);
        Assert.AreEqual(ws.Style.Font.Underline, ws.Cell("A1").Style.Font.Underline);
        Assert.IsNull(link.Container);
        Assert.IsFalse(ws.Hyperlinks.TryGet(ws.Cell("A1").Address, out _));
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
        Assert.IsFalse(wasDeleted);
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
        Assert.IsFalse(deleted);
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
        Assert.AreSame(link, foundLink);
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
            Assert.Throws<KeyNotFoundException>(() => ws.Hyperlinks.Get(addressWithoutLink));
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

        Assert.Throws<KeyNotFoundException>(() => ws1.Hyperlinks.Get(wrongSheetAddress));
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
        Assert.IsTrue(wasFound);
        Assert.AreSame(link, foundLink);
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
            Assert.IsFalse(ws.Hyperlinks.TryGet(addressWithoutLink, out _));
        }
    }
}
