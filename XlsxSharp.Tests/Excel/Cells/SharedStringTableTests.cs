using System;
using System.Text;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Cells;

[TestFixture]
public class SharedStringTableTests
{
    [Test]
    public void SameStringIsNotStoredTwice()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        IXLWorksheet ws2 = wb.AddWorksheet();
        string txt1 = "Hello";
        string txt2 = new StringBuilder("Hel").Append("lo").ToString();
        Assert.AreNotSame(txt1, txt2);

        ws1.Cell(1, 1).Value = txt1;
        ws2.Cell(1, 1).Value = txt2;

        Assert.AreSame(ws1.Cell(1, 1).Value.GetText(), ws2.Cell(1, 1).Value.GetText());
    }

    [Test]
    public void CanAccessTextThroughId()
    {
        SharedStringTable sst = new();
        int id = sst.IncreaseRef("test", false);
        Assert.AreEqual("test", sst[id]);
        Assert.AreEqual(1, sst.Count);
    }

    [Test]
    public void TextsWithoutReferenceAreRemoved()
    {
        SharedStringTable sst = new();
        int id = sst.IncreaseRef("test", false);
        sst.DecreaseRef(id);

        Assert.AreEqual(0, sst.Count);
        Assert.That(
            () => _ = sst[id],
            Throws.ArgumentException.With.Message.EqualTo("Id 0 has no text.")
        );
    }

    [Test]
    public void TextReferencedByMultipleThingsIsNotFreedUntilAllAreRelease()
    {
        const string text = "test";
        SharedStringTable sst = new();
        int id = sst.IncreaseRef(text, false);

        sst.IncreaseRef(text, false);
        Assert.AreEqual(text, sst[id]);
        Assert.AreEqual(1, sst.Count);

        sst.DecreaseRef(id);
        Assert.AreEqual(text, sst[id]);
        Assert.AreEqual(1, sst.Count);

        sst.IncreaseRef(text, false);
        Assert.AreEqual(text, sst[id]);
        Assert.AreEqual(1, sst.Count);

        sst.DecreaseRef(id);
        Assert.AreEqual(text, sst[id]);
        Assert.AreEqual(1, sst.Count);

        sst.DecreaseRef(id);
        Assert.Throws<ArgumentException>(() => _ = sst[id]);
    }

    [Test]
    public void FreedIdCanBeReusedForDifferentText()
    {
        SharedStringTable sst = new();
        sst.IncreaseRef("zero", false);
        int originalId = sst.IncreaseRef("original", false);
        int laterId = sst.IncreaseRef("two", false);

        Assert.That(laterId, Is.GreaterThan(originalId));

        sst.DecreaseRef(originalId);
        Assert.Throws<ArgumentException>(() => _ = sst[originalId]);

        int replacementId = sst.IncreaseRef("replacement", false);
        Assert.AreEqual(originalId, replacementId);
        Assert.AreEqual("replacement", sst[replacementId]);
    }

    [Test]
    public void DereferencingFreedIdThrows()
    {
        SharedStringTable sst = new();
        int id = sst.IncreaseRef("test", false);
        sst.DecreaseRef(id);
        Assert.Throws<InvalidOperationException>(() => sst.DecreaseRef(id));
    }

    [Test]
    public void StringItemWithoutTextIsLoadedAsEmptyText() =>
        // PR#2218: A text cell that references self-closed <si/> tag in SST is loaded without
        // an error and is loaded as type TEXT. Although it's not very common, empty string is
        // a valid value of a cell.
        TestHelper.LoadAndAssert(
            (_, ws) =>
            {
                // Check that type is a empty string, just like in Excel.
                Assert.AreEqual(2, ws.Evaluate("TYPE(B2)"));
                Assert.IsEmpty(ws.Cell("B2").GetText());
            },
            @"Other\Cells\EmptySi.xlsx"
        );

    [Test]
    public void EmptyTextIsWrittenAndLoadedToSst() =>
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = "Empty text cell (B1):";
                ws.Cell("B1").Value = string.Empty;

                ws.Cell("A2").Value = "Empty rich text";
                ws.Cell("B2").CreateRichText().AddText(string.Empty);
            },
            (_, ws) =>
            {
                Assert.AreEqual("", ws.Cell("B1").CachedValue);
                Assert.AreEqual("", ws.Cell("B2").GetRichText().Text);
            },
            @"Other\Cells\EmptyText.xlsx"
        );
}
