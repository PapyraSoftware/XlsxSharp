#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public enum XLCFIconSetOperator
{
    GreaterThan,
    EqualOrGreaterThan,
}

public interface IXLCFIconSet
{
    public IXLCFIconSet AddValue(
        XLCFIconSetOperator setOperator,
        string value,
        XLCFContentType type
    );
    public IXLCFIconSet AddValue(
        XLCFIconSetOperator setOperator,
        double value,
        XLCFContentType type
    );
}
