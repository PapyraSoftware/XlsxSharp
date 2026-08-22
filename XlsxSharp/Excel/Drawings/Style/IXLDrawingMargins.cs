namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingMargins
{
    public bool Automatic { get; set; }

    /// <summary>
    /// Left margin in inches.
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// Right margin in inches.
    /// </summary>
    public double Right { get; set; }

    /// <summary>
    /// Top margin in inches.
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// Bottom margin in inches.
    /// </summary>
    public double Bottom { get; set; }

    /// <summary>
    /// Set <see cref="Left"/>, <see cref="Top"/>, <see cref="Right"/>, <see cref="Bottom"/> margins at once.
    /// </summary>
    public double All { set; }

    public IXLDrawingStyle SetAutomatic();
    public IXLDrawingStyle SetAutomatic(bool value);
    public IXLDrawingStyle SetLeft(double value);
    public IXLDrawingStyle SetRight(double value);
    public IXLDrawingStyle SetTop(double value);
    public IXLDrawingStyle SetBottom(double value);
    public IXLDrawingStyle SetAll(double value);
}
