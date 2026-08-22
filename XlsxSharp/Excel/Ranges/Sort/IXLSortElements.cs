#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.Sort;

public interface IXLSortElements : IEnumerable<IXLSortElement>
{
    public void Add(int elementNumber);
    public void Add(int elementNumber, XLSortOrder sortOrder);
    public void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks);
    public void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase);

    public void Add(string elementNumber);
    public void Add(string elementNumber, XLSortOrder sortOrder);
    public void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks);
    public void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase);

    public void Clear();

    public void Remove(int elementNumber);
}
