#nullable disable

namespace XlsxSharp.Excel.CustomProperties;

public enum XLCustomPropertyType
{
    Text,
    Number,
    Date,
    Boolean,
}

public interface IXLCustomProperty
{
    public string Name { get; set; }
    public XLCustomPropertyType Type { get; }
    public object Value { get; set; }
    public T GetValue<T>();
}
