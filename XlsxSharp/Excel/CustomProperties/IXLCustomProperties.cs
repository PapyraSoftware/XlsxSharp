#nullable disable

using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel.CustomProperties;

public interface IXLCustomProperties : IEnumerable<IXLCustomProperty>
{
    void Add(IXLCustomProperty customProperty);
    void Add<T>(String name, T value);
    void Delete(String name);
    IXLCustomProperty CustomProperty(String name);
}
