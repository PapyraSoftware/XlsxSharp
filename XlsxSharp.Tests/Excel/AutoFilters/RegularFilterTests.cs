using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.AutoFilters;

[TestFixture]
public class RegularFilterTests
{
    [Test]
    public void DateTimeGroupingAndRegularValuesCanBeUsedTogether()
    {
        // OpenXML SDK validator considers filter and dateTimeGroup filter elements together to
        // be an error, but it isn't (XSD allows and Excel reads). Therefore, disable
        // validation for the test.
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
                            new DateTime(2015, 7, 25),
                            new DateTime(2015, 8, 25),
                        }
                    )
                    .SetAutoFilter();
                autoFilter
                    .Column(1)
                    .AddFilter(1)
                    .AddDateGroupFilter(new DateTime(2015, 8, 1), XLDateTimeGrouping.Month);
            },
            (_, ws) =>
            {
                ws.AutoFilter.Reapply();
                IEnumerable<bool> dataVisibility = ws.Rows("2:5").Select(row => !row.IsHidden);
                CollectionAssert.AreEqual(new[] { true, false, false, true }, dataVisibility);
            },
            false
        );
    }

    [Test]
    [SetCulture("cs-CZ")]
    public void RegularNumberValueIsComparedAsTextAgainstFormattedText()
    {
        new AutoFilterTester(f => f.AddFilter(1.5))
            .Add(1.5, true)
            .Add("1.5", false)
            .Add("1,5", true)
            .Add("1,50", false)
            .Add(
                1.5,
                nf => nf.SetNumberFormatId((int)XLPredefinedFormat.Number.PercentPrecision2),
                false
            )
            .Add(700, nf => nf.SetFormat("\"1,5\""), true)
            .AssertVisibility();
    }

    [Test]
    [SetCulture("cs-CZ")]
    public void RegularLogicalValueIsComparedAsTextAgainstFormattedText()
    {
        new AutoFilterTester(f => f.AddFilter(false))
            .Add(false, true)
            .Add(0, false)
            .Add("FALSE", true)
            .Add("TRUE", false)
            .Add(true, false)
            .Add(77, nf => nf.SetFormat("\"FALSE\""), true)
            .AssertVisibility();
    }

    [Test]
    [SetCulture("cs-CZ")]
    public void RegularErrorValueIsComparedAsTextAgainstFormattedText()
    {
        new AutoFilterTester(f => f.AddFilter("#VALUE!"))
            .Add(XLError.IncompatibleValue, true)
            .Add(2, false)
            .Add("#VALUE!", true)
            .AssertVisibility();
    }

    [Test]
    public void PatternIsNotInterpretedAsWildcard()
    {
        new AutoFilterTester(f => f.AddFilter("A*"))
            .Add("A*", true)
            .Add("A", false)
            .Add("A something", false)
            .AssertVisibility();
    }
}
