using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// A named style for a pivot table.
/// </summary>
internal class XLPivotTableStyle
{
    private readonly Dictionary<XLPivotStyleRegionValues, XLDxfValue> _regionFormats = new();

    public XLPivotTableStyle(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            throw new ArgumentException("Name can't be empty.");
        }

        this.Name = name;
    }

    internal string Name { get; }

    internal IReadOnlyDictionary<XLPivotStyleRegionValues, XLDxfValue> RegionFormats =>
        this._regionFormats;

    /// <summary>
    /// A band size (i.e. how many rows does a stripe have) for a <see cref="XLPivotStyleRegionValues.FirstRowStripe"/>.
    /// </summary>
    internal int RowStripe1BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many rows does a stripe have) for a <see cref="XLPivotStyleRegionValues.SecondRowStripe"/>.
    /// </summary>
    internal int RowStripe2BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many column does a stripe have) for a <see cref="XLPivotStyleRegionValues.FirstColumnStripe"/>.
    /// </summary>
    internal int ColumnStripe1BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many column does a stripe have) for a <see cref="XLPivotStyleRegionValues.SecondColumnStripe"/>.
    /// </summary>
    internal int ColumnStripe2BandSize { get; private set; } = 1;

    internal void SetRegionFormat(XLPivotStyleRegionValues region, XLDxfValue dxf, int bandSize = 1)
    {
        // Keep values excel compatible.
        if (bandSize is < 0 or > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(bandSize));
        }

        // Band size has only meaning for stripe regions
        switch (region)
        {
            case XLPivotStyleRegionValues.FirstRowStripe:
                this.RowStripe1BandSize = bandSize;
                break;
            case XLPivotStyleRegionValues.SecondRowStripe:
                this.RowStripe2BandSize = bandSize;
                break;
            case XLPivotStyleRegionValues.FirstColumnStripe:
                this.ColumnStripe1BandSize = bandSize;
                break;
            case XLPivotStyleRegionValues.SecondColumnStripe:
                this.ColumnStripe2BandSize = bandSize;
                break;
        }

        this._regionFormats[region] = dxf;
    }
}
