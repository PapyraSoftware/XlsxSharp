#nullable disable

using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.PivotStyleFormats;

internal class XLPivotTableAxisFieldStyleFormats : IXLPivotFieldStyleFormats
{
    private readonly XLPivotTable _pivotTable;
    private readonly XLPivotTableAxisField _axisField;

    public XLPivotTableAxisFieldStyleFormats(
        XLPivotTable pivotTable,
        XLPivotTableAxisField axisField
    )
    {
        this._pivotTable = pivotTable;
        this._axisField = axisField;
    }

    #region IXLPivotFieldStyleFormats

    public IXLPivotValueStyleFormat DataValuesFormat =>
        new XLPivotValueStyleFormat(this._pivotTable, this._axisField.Offset);

    public IXLPivotStyleFormat Header
    {
        get
        {
            /*
             * <x:pivotArea field="4"
             *  type="button"
             *  axis="axisCol"
             *  fieldPosition="0"/>
             *
             * The area must have field position and axis, otherwise the style is not correctly
             * displayed.
             */
            // If table is not compact, each field has it's own header and thus pivot area must
            // contain correct position of the field on the axis. If table is compact, there is
            // only one header and its position is first in axis, because it's the only one.
            int fieldPosition = this._pivotTable.Compact ? 0 : this._axisField.Position;
            XLPivotAxis fieldAxis = this._axisField.Axis;
            XLPivotArea headerArea = new()
            {
                Field = this._axisField.Offset,
                Type = XLPivotAreaType.Button,
                Axis = fieldAxis,
                FieldPosition = (uint)fieldPosition,
            };

            return new XLPivotStyleFormat(
                this._pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, headerArea),
                () => headerArea
            );
        }
    }

    public IXLPivotStyleFormat Label
    {
        get
        {
            /* <x:pivotArea type="normal"
             *              dataOnly="0"
             *              labelOnly="1">
             *     <x:references count="1">
             *         <x:reference field="4"/>
             *	   </x:references>
             * </x:pivotArea>
             */
            XLPivotArea labelArea = new() { DataOnly = false, LabelOnly = true };
            labelArea.AddReference(new XLPivotReference { Field = (uint)this._axisField.Offset });

            return new XLPivotStyleFormat(
                this._pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, labelArea),
                () => labelArea
            );
        }
    }

    public IXLPivotStyleFormat Subtotal
    {
        get
        {
            /* <pivotArea outline="0">
             *   <references count="1">
             *     <reference field="0"
             *                count="0"
             *                defaultSubtotal="1"/>
             *   </references>
             * </pivotArea>
             */
            // Subtotal fields in reference can't mix default and custom subtotals. It always must
            // reference only one type. Excel doesn't select correct area if they are mixed.
            // The outline flag has weird behavior, but is required for subtotals of last field in
            // an axis with multiple fields (i.e. subtotals are displayed at the bottom).
            HashSet<XLSubtotalFunction> subtotals = [.. this._axisField.Subtotals];
            XLPivotArea subtotalArea = new() { Outline = false };
            subtotalArea.AddReference(
                new XLPivotReference
                {
                    Field = unchecked((uint)this._axisField.Offset),
                    Subtotals = subtotals,
                }
            );

            return new XLPivotStyleFormat(
                this._pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, subtotalArea),
                () => subtotalArea
            );
        }
    }

    #endregion
}
