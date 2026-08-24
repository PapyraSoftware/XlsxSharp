using XlsxSharp.Excel.PivotValues;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.PivotStyleFormats;

internal class XLPivotValueStyleFormat : XLPivotStyleFormatBase, IXLPivotValueStyleFormat
{
    /// <summary>
    /// A list of references that specify which data cells will be styled.
    /// A data cell will be styled, if it lies on all referenced fields.
    /// The term "lie on" means that either column or a row of data cell
    /// intersects a label cell of referenced field.
    /// </summary>
    private readonly List<FieldReference> _fieldReferences = [];

    public XLPivotValueStyleFormat(XLPivotTable pivotTable, FieldIndex fieldIndex)
        : base(pivotTable) => this._fieldReferences.Add(new FieldReference(fieldIndex));

    #region IXLPivotValueStyleFormat members

    public IXLPivotValueStyleFormat AndWith(IXLPivotField field)
    {
        this._fieldReferences.Add(new FieldReference(field.Offset));
        return this;
    }

    public IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<XLCellValue>? predicate)
    {
        FieldIndex fieldIndex = field.Offset;
        if (fieldIndex.IsDataField)
        {
            throw new ArgumentException("Field is a 'data' field.", nameof(field));
        }

        if (predicate is null)
        {
            return this.AndWith(field);
        }

        XLPivotTableField pivotField = this.PivotTable.PivotFields[fieldIndex];
        List<uint> filteredItems = pivotField
            .GetAllItems(predicate)
            .WhereNotNull(fieldItem => fieldItem.ItemIndex)
            .Select(itemIndex => (uint)itemIndex)
            .ToList();

        this._fieldReferences.Add(new FieldReference(field.Offset, filteredItems));
        return this;
    }

    public IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField)
    {
        int valuesIndex = this.PivotTable.DataFields.IndexOf(valueField);
        if (valuesIndex == -1)
        {
            throw new ArgumentOutOfRangeException(
                $"Field '{valueField.CustomName}' is not among value fields of the pivot table."
            );
        }

        this._fieldReferences.Add(
            new FieldReference(FieldIndex.DataField, new[] { (uint)valuesIndex })
        );
        return this;
    }

    #endregion IXLPivotValueStyleFormat members

    internal override XLPivotArea GetCurrentArea()
    {
        XLPivotArea area = new();
        foreach (FieldReference fieldReference in this._fieldReferences)
        {
            XLPivotReference reference = new()
            {
                Field = unchecked((uint?)fieldReference.FieldIndex.Value),
            };
            if (fieldReference.Items is not null)
            {
                foreach (uint item in fieldReference.Items)
                {
                    reference.AddFieldItem(item);
                }
            }

            area.AddReference(reference);
        }

        return area;
    }

    internal override bool Filter(XLPivotArea area)
    {
        XLPivotArea currentArea = this.GetCurrentArea();
        return XLPivotAreaComparer.Instance.Equals(area, currentArea);
    }

    private record FieldReference(FieldIndex FieldIndex, IReadOnlyList<uint>? Items = null);
}
