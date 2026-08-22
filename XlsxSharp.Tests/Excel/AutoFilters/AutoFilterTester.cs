using System;
using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.AutoFilters;

internal class AutoFilterTester
{
    private readonly Action<IXLFilterColumn> setFilter;
    private readonly List<(
        XLCellValue Value,
        Action<IXLStyle> SetStyle,
        bool ExpectedVisibility
    )> values = [];

    internal AutoFilterTester(Action<IXLFilterColumn> setFilter) => this.setFilter = setFilter;

    internal AutoFilterTester Add(XLCellValue value, bool shouldBeVisible) =>
        this.Add(value, static (IXLStyle _) => { }, shouldBeVisible);

    internal AutoFilterTester Add(
        XLCellValue value,
        Action<IXLNumberFormat> setNumberFormat,
        bool shouldBeVisible
    )
    {
        this.values.Add((value, s => setNumberFormat(s.NumberFormat), shouldBeVisible));
        return this;
    }

    internal AutoFilterTester Add(
        XLCellValue value,
        Action<IXLStyle> setStyle,
        bool shouldBeVisible
    )
    {
        this.values.Add((value, setStyle, shouldBeVisible));
        return this;
    }

    internal AutoFilterTester AddTrue(params XLCellValue[] values)
    {
        foreach (XLCellValue value in values)
        {
            this.Add(value, true);
        }

        return this;
    }

    internal AutoFilterTester AddFalse(params XLCellValue[] values)
    {
        foreach (XLCellValue value in values)
        {
            this.Add(value, false);
        }

        return this;
    }

    internal void AssertVisibility()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Data";
        for (int i = 0; i < this.values.Count; ++i)
        {
            IXLCell cell = ws.Cell(i + 2, 1);
            cell.Value = this.values[i].Value;
            this.values[i].SetStyle(cell.Style);
        }

        IXLAutoFilter autoFilter = ws.Range(1, 1, this.values.Count + 1, 1).SetAutoFilter();
        this.setFilter(autoFilter.Column(1));

        for (int i = 0; i < this.values.Count; ++i)
        {
            int row = i + 2;
            XLCellValue value = ws.Cell(row, 1).CachedValue;
            string formattedString = ((XLCell)ws.Cell(row, 1)).GetFormattedString(value);
            bool actualVisible = !ws.Row(row).IsHidden;
            bool expectedVisibility = this.values[i].ExpectedVisibility;
            Assert.AreEqual(
                expectedVisibility,
                actualVisible,
                $"Visibility differs at index {i} for value {value} (formatted '{formattedString}')"
            );
        }
    }
}
