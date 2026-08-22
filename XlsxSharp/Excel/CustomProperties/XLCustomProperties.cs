#nullable disable

using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel.CustomProperties;

internal class XLCustomProperties : IXLCustomProperties, IEnumerable<IXLCustomProperty>
{
    XLWorkbook workbook;

    public XLCustomProperties(XLWorkbook workbook)
    {
        this.workbook = workbook;
    }

    private Dictionary<String, IXLCustomProperty> customProperties = new();

    public void Add(IXLCustomProperty customProperty)
    {
        this.customProperties.Add(customProperty.Name, customProperty);
    }

    public void Add<T>(String name, T value)
    {
        XLCustomProperty cp = new(this.workbook) { Name = name, Value = value };
        this.Add(cp);
    }

    public void Delete(String name)
    {
        this.customProperties.Remove(name);
    }

    public IXLCustomProperty CustomProperty(String name)
    {
        return this.customProperties[name];
    }

    public IEnumerator<IXLCustomProperty> GetEnumerator()
    {
        return this.customProperties.Values.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}
