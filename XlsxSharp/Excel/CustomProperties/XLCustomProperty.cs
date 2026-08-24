#nullable disable

namespace XlsxSharp.Excel.CustomProperties;

internal class XLCustomProperty : IXLCustomProperty
{
    private readonly XLWorkbook _workbook;

    private string name;

    public XLCustomProperty(XLWorkbook workbook) => this._workbook = workbook;

    #region IXLCustomProperty Members

    public string Name
    {
        get => this.name;
        set
        {
            if (this.name == value)
            {
                return;
            }

            if (this._workbook.CustomProperties.Any(t => t.Name == value))
            {
                throw new ArgumentException(
                    string.Format(
                        "This workbook already contains a custom property named '{0}'",
                        value
                    )
                );
            }

            this.name = value;
        }
    }

    public XLCustomPropertyType Type
    {
        get
        {
            if (this.Value is DateTime)
            {
                return XLCustomPropertyType.Date;
            }

            if (this.Value is bool)
            {
                return XLCustomPropertyType.Boolean;
            }

            if (double.TryParse(this.Value.ToString(), out double dTest))
            {
                return XLCustomPropertyType.Number;
            }

            return XLCustomPropertyType.Text;
        }
    }

    public object Value { get; set; }

    public T GetValue<T>() => (T)Convert.ChangeType(this.Value, typeof(T));

    #endregion
}
