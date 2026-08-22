namespace XlsxSharp.Excel.PivotStyleFormats;

/// <summary>
/// Interface to change the style of a <see cref="IXLPivotField"/> or its parts.
/// </summary>
public interface IXLPivotFieldStyleFormats
{
    /// <summary>
    /// Pivot table style of the field values displayed in the data area of the pivot table.
    /// </summary>
    public IXLPivotValueStyleFormat DataValuesFormat { get; }

    /// <summary>
    /// Get the style of the pivot field header. The head usually contains a name of the field.
    /// In some layouts, header is not individually displayed (e.g. compact), while in others
    /// it is (e.g. tabular).
    /// </summary>
    public IXLPivotStyleFormat Header { get; }

    /// <summary>
    /// Get the style of the pivot field label values on horizontal or vertical axis.
    /// </summary>
    public IXLPivotStyleFormat Label { get; }

    public IXLPivotStyleFormat Subtotal { get; }
}
