// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections;
using System.Collections.Generic;

namespace XlsxSharp.Excel.PivotStyleFormats;

/// <summary>
/// An API for grand totals from <see cref="XLPivotTableStyleFormats"/>.
/// </summary>
internal class XLPivotStyleFormats : IXLPivotStyleFormats
{
    private readonly XLPivotTable _pivotTable;
    private readonly bool _isRowGrand;

    internal XLPivotStyleFormats(XLPivotTable pivotTable, bool isRowGrand)
    {
        this._pivotTable = pivotTable;
        this._isRowGrand = isRowGrand;
    }

    #region IXLPivotStyleFormats members

    public IXLPivotStyleFormat ForElement(XLPivotStyleFormatElement element)
    {
        if (element == XLPivotStyleFormatElement.None)
        {
            throw new ArgumentException(
                "Choose an enum value that represents an element",
                nameof(element)
            );
        }

        return this.GetPivotStyleFormatFor(element);
    }

    public IEnumerator<IXLPivotStyleFormat> GetEnumerator()
    {
        XLPivotStyleFormatElement[] elements =
        [
            XLPivotStyleFormatElement.Label,
            XLPivotStyleFormatElement.Data,
            XLPivotStyleFormatElement.All,
        ];

        foreach (XLPivotStyleFormatElement element in elements)
        {
            foreach (XLPivotFormat format in this._pivotTable.Formats)
            {
                if (this.AreaBelongsToGrandTotal(format.PivotArea, element))
                {
                    // Each pivot style format modifies all formats, so return only once per element.
                    yield return this.GetPivotStyleFormatFor(element);
                    break;
                }
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    #endregion IXLPivotStyleFormats members

    private XLPivotStyleFormat GetPivotStyleFormatFor(XLPivotStyleFormatElement element)
    {
        return new XLPivotStyleFormat(this._pivotTable, FilterElement, ElementFactory)
        {
            AppliesTo = element,
        };

        bool FilterElement(XLPivotArea pivotArea) =>
            this.AreaBelongsToGrandTotal(pivotArea, element);
        XLPivotArea ElementFactory() => this.CreateGrandArea(element);
    }

    private bool AreaBelongsToGrandTotal(XLPivotArea area, XLPivotStyleFormatElement element) =>
        area.References.Count == 0
        && area.Field is null
        && area.Type == XLPivotAreaType.Normal
        && area.DataOnly == (element == XLPivotStyleFormatElement.Data)
        && area.LabelOnly == (element == XLPivotStyleFormatElement.Label)
        && area.GrandRow == this._isRowGrand
        && area.GrandCol == !this._isRowGrand
        && area.CacheIndex == false
        && area.Offset is null
        && !area.CollapsedLevelsAreSubtotals
        && area.Axis is null
        && area.FieldPosition is null;

    private XLPivotArea CreateGrandArea(XLPivotStyleFormatElement element) =>
        new()
        {
            DataOnly = (element == XLPivotStyleFormatElement.Data),
            LabelOnly = (element == XLPivotStyleFormatElement.Label),
            GrandRow = this._isRowGrand,
            GrandCol = !this._isRowGrand,
        };
}
