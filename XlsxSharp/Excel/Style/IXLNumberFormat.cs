namespace XlsxSharp.Excel;

public interface IXLNumberFormat : IXLNumberFormatBase, IEquatable<IXLNumberFormatBase>
{
    /// <summary>
    /// Sets a predefined number format. Use members of <see cref="XLPredefinedFormat.General"/>,
    /// <see cref="XLPredefinedFormat.Number"/> or <see cref="XLPredefinedFormat.DateTime"/>
    /// casted to <c>int</c>.
    /// <example>
    /// <code>
    ///   cell.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2);
    /// </code>
    /// </example>
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The passed value is not a predefined format.</exception>
    public IXLStyle SetNumberFormatId(int value);

    /// <summary>
    /// Sets the number format.
    /// </summary>
    public IXLStyle SetFormat(string value);
}
