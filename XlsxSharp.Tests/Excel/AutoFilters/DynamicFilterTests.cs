using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.AutoFilters;

public class DynamicFilterTests
{
    [Test]
    public void AverageFilterIsInitializedAfterLoad() =>
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                IXLAutoFilter autoFilter = ws.Cell("A1")
                    .InsertData(
                        new object[]
                        {
                            "Data",
                            1,
                            2,
                            3,
                            4,
                            5,
                            10, // avg. 4.16
                        }
                    )
                    .SetAutoFilter();
                autoFilter.Column(1).AboveAverage();
            },
            (_, ws) =>
            {
                ws.AutoFilter.Reapply();
                IEnumerable<bool> filterResult = ws.Rows("2:7").Select(row => !row.IsHidden);
                CollectionAssert.AreEqual(
                    new[] { false, false, false, false, true, true },
                    filterResult
                );
            }
        );

    [Test]
    public void BelowAverageTakesValuesUnderAvgValue() =>
        // The average 2 is not included.
        new AutoFilterTester(f => f.BelowAverage())
            .AddTrue(1)
            .AddFalse(2, 3)
            .AssertVisibility();

    [Test]
    public void AboveAverageTakesValuesOverAvgValue() =>
        new AutoFilterTester(f => f.AboveAverage()).AddTrue(3).AddFalse(2, 1).AssertVisibility();

    [Test]
    public void AverageIgnoresNonUnifiedNumbers() =>
        new AutoFilterTester(f => f.BelowAverage())
            .AddTrue(new DateTime(1900, 1, 1)) // Serial date time 1
            .AddFalse(1.1)
            .AddFalse(1.2)
            .AddFalse(XLError.NoValueAvailable, true, false, "-100", "Text", Blank.Value)
            .AssertVisibility();

    [Test]
    public void AllRowsAreHiddenWhenColumnHasNoNumber() =>
        new AutoFilterTester(f => f.AboveAverage())
            .AddFalse(Blank.Value, true, false, "-100", "Text", XLError.NoValueAvailable)
            .AssertVisibility();
}
