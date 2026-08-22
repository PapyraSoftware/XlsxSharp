#nullable disable

using System;
using System.Linq;

namespace XlsxSharp.Excel.CustomProperties;

internal class XLCustomProperty : IXLCustomProperty
{
    private readonly XLWorkbook _workbook;

    private String name;

    public XLCustomProperty(XLWorkbook workbook)
    {
        this._workbook = workbook;
    }

    #region IXLCustomProperty Members

    public String Name
    {
        get { return this.name; }
        set
        {
            if (this.name == value)
            {
                return;
            }

            if (this._workbook.CustomProperties.Any(t => t.Name == value))
            {
                throw new ArgumentException(
                    String.Format(
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

            if (this.Value is Boolean)
            {
                return XLCustomPropertyType.Boolean;
            }

            if (Double.TryParse(this.Value.ToString(), out Double dTest))
            {
                return XLCustomPropertyType.Number;
            }

            return XLCustomPropertyType.Text;
        }
    }

    public Object Value { get; set; }

    public T GetValue<T>()
    {
        return (T)Convert.ChangeType(this.Value, typeof(T));
    }

    #endregion
}
