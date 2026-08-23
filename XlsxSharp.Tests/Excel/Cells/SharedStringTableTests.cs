using System;
using System.Text;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Cells;

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
        ClassicAssert.AreNotSame(txt1, txt2);

        ws1.Cell(1, 1).Value = txt1;
        ws2.Cell(1, 1).Value = txt2;

        ClassicAssert.AreSame(ws1.Cell(1, 1).Value.GetText(), ws2.Cell(1, 1).Value.GetText());
    }

    [Test]
    public void CanAccessTextThroughId()
    {
        SharedStringTable sst = new();
        int id = sst.IncreaseRef("test", false);
        ClassicAssert.AreEqual("test", sst[id]);
        ClassicAssert.AreEqual(1, sst.Count);
    }

    [Test]
    public void TextsWithoutReferenceAreRemoved()
    {
        SharedStringTable sst = new();
        int id = sst.IncreaseRef("test", false);
        sst.DecreaseRef(id);

        ClassicAssert.AreEqual(0, sst.Count);
        ArgumentException ex = ClassicAssert.Throws<ArgumentException>(() => _ = sst[id]);
        ClassicAssert.AreEqual("Id 0 has no text.", ex.Message);
    }

    [Test]
    public void TextReferencedByMultipleThingsIsNotFreedUntilAllAreRelease()
    {
        const string text = "test";
        SharedStringTable sst = new();
        int id = sst.IncreaseRef(text, false);

        sst.IncreaseRef(text, false);
        ClassicAssert.AreEqual(text, sst[id]);
        ClassicAssert.AreEqual(1, sst.Count);

        sst.DecreaseRef(id);
        ClassicAssert.AreEqual(text, sst[id]);
        ClassicAssert.AreEqual(1, sst.Count);

        sst.IncreaseRef(text, false);
        ClassicAssert.AreEqual(text, sst[id]);
        ClassicAssert.AreEqual(1, sst.Count);

        sst.DecreaseRef(id);
        ClassicAssert.AreEqual(text, sst[id]);
        ClassicAssert.AreEqual(1, sst.Count);

        sst.DecreaseRef(id);
        ClassicAssert.Throws<ArgumentException>(() => _ = sst[id]);
    }

    [Test]
    public void FreedIdCanBeReusedForDifferentText()
    {
        SharedStringTable sst = new();
        sst.IncreaseRef("zero", false);
        int originalId = sst.IncreaseRef("original", false);
        int laterId = sst.IncreaseRef("two", false);

        ClassicAssert.Greater(laterId, originalId);

        sst.DecreaseRef(originalId);
        ClassicAssert.Throws<ArgumentException>(() => _ = sst[originalId]);

        int replacementId = sst.IncreaseRef("replacement", false);
        ClassicAssert.AreEqual(originalId, replacementId);
        ClassicAssert.AreEqual("replacement", sst[replacementId]);
    }

    [Test]
    public void DereferencingFreedIdThrows()
    {
        SharedStringTable sst = new();
        int id = sst.IncreaseRef("test", false);
        sst.DecreaseRef(id);
        ClassicAssert.Throws<InvalidOperationException>(() => sst.DecreaseRef(id));
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
                ClassicAssert.AreEqual(2, ws.Evaluate("TYPE(B2)"));
                ClassicAssert.IsEmpty(ws.Cell("B2").GetText());
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
                ClassicAssert.AreEqual("", ws.Cell("B1").CachedValue);
                ClassicAssert.AreEqual("", ws.Cell("B2").GetRichText().Text);
            },
            @"Other\Cells\EmptyText.xlsx"
        );
}
