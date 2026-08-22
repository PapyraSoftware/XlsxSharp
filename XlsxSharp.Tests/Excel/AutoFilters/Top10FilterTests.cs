using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.AutoFilters;

[TestFixture]
public class Top10FilterTests
{
    [Test]
    public void Top10FilterIsInitializedAfterLoad()
    {
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                IXLAutoFilter autoFilter = ws.Cell("A1")
                    .InsertData(new object[] { "Data", 4, 4, 1, 3, 2, 5 })
                    .SetAutoFilter();
                autoFilter.Column(1).Top(3);
            },
            (_, ws) =>
            {
                ws.AutoFilter.Reapply();
                IEnumerable<bool> filterResult = ws.Rows("2:7").Select(row => !row.IsHidden);
                CollectionAssert.AreEqual(
                    new[] { true, true, false, false, false, true },
                    filterResult
                );
            }
        );
    }

    [Test]
    public void TopItemsFilterExcludesNonUnifiedNumbers()
    {
        // Sort and then use cutoff value, it's 4 here and then take all values >= cutoff.
        new AutoFilterTester(f => f.Top(1))
            .AddTrue(new DateTime(1900, 2, 10))
            .AddFalse(11, 10)
            .AddFalse("-1000", "Text", Blank.Value, true, false, XLError.IncompatibleValue)
            .AssertVisibility();
    }

    [Test]
    public void BottomItemsFilterExcludesNonUnifiedNumbers()
    {
        new AutoFilterTester(f => f.Bottom(1))
            .AddTrue(new DateTime(1900, 1, 1))
            .AddFalse(2, 3)
            .AddFalse("-1000", "Text", Blank.Value, true, false, XLError.IncompatibleValue)
            .AssertVisibility();
    }

    [Test]
    public void TopItemsFilterDeterminesTopItemsByDeterminingCutOffValue()
    {
        // Sort and then use cutoff value, it's 4 here and then take all values <= cutoff.
        new AutoFilterTester(f => f.Top(2))
            .AddTrue(5, 4, 4, 4)
            .AddFalse(3, 2, 1)
            .AssertVisibility();

        // Cutoff is 5 here.
        new AutoFilterTester(f => f.Top(2))
            .AddTrue(5, 5)
            .AddFalse(4, 4, 4, 3, 2, 1)
            .AssertVisibility();
    }

    [Test]
    public void BottomItemsFilterDeterminesTopItemsByDeterminingCutOffValue()
    {
        // Cutoff is 2
        new AutoFilterTester(f => f.Bottom(2))
            .AddTrue(1, 2, 2, 2)
            .AddFalse(3, 4, 5)
            .AssertVisibility();

        // Cutoff is 5
        new AutoFilterTester(f => f.Bottom(2))
            .AddTrue(1, 1)
            .AddFalse(2, 2, 2, 3, 4, 5)
            .AssertVisibility();
    }

    [Test]
    public void TopPercentsUsesInclusivePercentValue()
    {
        // Autofilter doesn't include value 750, which is at 75%, i.e. right at the border.
        new AutoFilterTester(f => f.Top(25, XLTopBottomType.Percent))
            .AddFalse([.. Enumerable.Range(1, 750).Select<int, XLCellValue>(x => x)])
            .AddTrue([.. Enumerable.Range(751, 250).Select<int, XLCellValue>(x => x)])
            .AssertVisibility();
    }

    [Test]
    public void BottomPercentsUsesInclusivePercentValue()
    {
        new AutoFilterTester(f => f.Bottom(25, XLTopBottomType.Percent))
            .AddTrue([.. Enumerable.Range(1, 250).Select<int, XLCellValue>(x => x)])
            .AddFalse([.. Enumerable.Range(251, 750).Select<int, XLCellValue>(x => x)])
            .AssertVisibility();
    }

    [Test]
    public void TopPercentsAlwaysHasAtLeastOneItem()
    {
        // Top 1% takes one item that is 33% of all items.
        new AutoFilterTester(f => f.Top(1, XLTopBottomType.Percent))
            .AddTrue(3)
            .AddFalse(2, 1)
            .AssertVisibility();
    }

    [Test]
    public void BottomPercentsAlwaysHasAtLeastOneItem()
    {
        new AutoFilterTester(f => f.Bottom(1, XLTopBottomType.Percent))
            .AddTrue(1)
            .AddFalse(2, 3)
            .AssertVisibility();
    }

    [TestCase(0, true)]
    [TestCase(501, true)]
    [TestCase(0, false)]
    [TestCase(501, false)]
    public void TopAndBottomFilterValueMustBeBetween1And500(int value, bool top)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Data";
        ws.Cell("A2").Value = value;
        IXLAutoFilter autoFilter = ws.Range("A1:A2").SetAutoFilter();
        IXLFilterColumn filterColumn = autoFilter.Column(1);

        ArgumentOutOfRangeException ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
        {
            if (top)
            {
                filterColumn.Top(value);
            }
            else
            {
                filterColumn.Bottom(value);
            }
        })!;
        StringAssert.Contains("Value must be between 1 and 500.", ex.Message);
    }
}
