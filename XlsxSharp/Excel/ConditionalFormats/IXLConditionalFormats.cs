using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLConditionalFormats : IEnumerable<IXLConditionalFormat>
{
    public void Add(IXLConditionalFormat conditionalFormat);

    public void RemoveAll();

    public void Remove(Predicate<IXLConditionalFormat> predicate);
}
