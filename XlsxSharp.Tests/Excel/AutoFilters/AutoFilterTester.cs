using System;
using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.AutoFilters;

internal class AutoFilterTester
{
    private readonly Action<IXLFilterColumn> _setFilter;
    private readonly List<(
        XLCellValue Value,
        Action<IXLStyle> SetStyle,
        bool ExpectedVisibility
    )> _values = [];

    internal AutoFilterTester(Action<IXLFilterColumn> setFilter)
    {
        this._setFilter = setFilter;
    }

    internal AutoFilterTester Add(XLCellValue value, bool shouldBeVisible)
    {
        return this.Add(value, static (IXLStyle _) => { }, shouldBeVisible);
    }

    internal AutoFilterTester Add(
        XLCellValue value,
        Action<IXLNumberFormat> setNumberFormat,
        bool shouldBeVisible
    )
    {
        this._values.Add((value, s => setNumberFormat(s.NumberFormat), shouldBeVisible));
        return this;
    }

    internal AutoFilterTester Add(
        XLCellValue value,
        Action<IXLStyle> setStyle,
        bool shouldBeVisible
    )
    {
        this._values.Add((value, setStyle, shouldBeVisible));
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
        for (int i = 0; i < this._values.Count; ++i)
        {
            IXLCell cell = ws.Cell(i + 2, 1);
            cell.Value = this._values[i].Value;
            this._values[i].SetStyle(cell.Style);
        }

        IXLAutoFilter autoFilter = ws.Range(1, 1, this._values.Count + 1, 1).SetAutoFilter();
        this._setFilter(autoFilter.Column(1));

        for (int i = 0; i < this._values.Count; ++i)
        {
            int row = i + 2;
            XLCellValue value = ws.Cell(row, 1).CachedValue;
            string formattedString = ((XLCell)ws.Cell(row, 1)).GetFormattedString(value);
            bool actualVisible = !ws.Row(row).IsHidden;
            bool expectedVisibility = this._values[i].ExpectedVisibility;
            Assert.AreEqual(
                expectedVisibility,
                actualVisible,
                $"Visibility differs at index {i} for value {value} (formatted '{formattedString}')"
            );
        }
    }
}
