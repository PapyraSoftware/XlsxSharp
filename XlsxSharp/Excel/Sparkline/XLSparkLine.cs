namespace XlsxSharp.Excel;

/// <summary>
/// An API object for a sparkline. It doesn't contain any data, only a link to the point of
/// the sparkline and does operations through <see cref="XLSparklineGroup"/>. It uses the
/// cell location as an anchor and if it is no longer valid, the group should throw.
/// </summary>
internal class XLSparkline : IXLSparkline
{
    private readonly XLSparklineGroup _sparklineGroup;
    private Point _location;

    internal XLSparkline(XLSparklineGroup sparklineGroup, Point location)
    {
        this._sparklineGroup = sparklineGroup;
        this._location = location;
    }

    public IXLCell Location
    {
        get => this._sparklineGroup.GetLocation(this._location);
        set => this.SetLocation(value);
    }

    public IXLRange? SourceData
    {
        get => this._sparklineGroup.GetSparklineSourceData(this._location);
        set => this.SetSourceData(value);
    }

    public IXLSparklineGroup SparklineGroup => this._sparklineGroup;

    public IXLSparkline SetLocation(IXLCell newLocation)
    {
        ArgumentNullException.ThrowIfNull(newLocation);

        if (newLocation.Worksheet != this.SparklineGroup.Worksheet)
        {
            throw new ArgumentException("Cannot move the sparkline to a different worksheet");
        }

        Point destination = Point.FromCell(newLocation);
        this._sparklineGroup.MoveSparkline(this._location, destination);
        this._location = destination;
        return this;
    }

    public IXLSparkline SetSourceData(IXLRange? sourceDataRange)
    {
        this._sparklineGroup.SetSparklineSourceData(this._location, sourceDataRange);
        return this;
    }
}
