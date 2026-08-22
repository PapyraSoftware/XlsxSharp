#nullable disable

using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel.Sort;

internal class XLSortElements : IXLSortElements
{
    List<IXLSortElement> elements = [];

    public void Add(Int32 elementNumber)
    {
        this.Add(elementNumber, XLSortOrder.Ascending);
    }

    public void Add(Int32 elementNumber, XLSortOrder sortOrder)
    {
        this.Add(elementNumber, sortOrder, true);
    }

    public void Add(Int32 elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks)
    {
        this.Add(elementNumber, sortOrder, ignoreBlanks, false);
    }

    public void Add(
        Int32 elementNumber,
        XLSortOrder sortOrder,
        Boolean ignoreBlanks,
        Boolean matchCase
    )
    {
        this.elements.Add(new XLSortElement(elementNumber, sortOrder, ignoreBlanks, matchCase));
    }

    public void Add(String elementNumber)
    {
        this.Add(elementNumber, XLSortOrder.Ascending);
    }

    public void Add(String elementNumber, XLSortOrder sortOrder)
    {
        this.Add(elementNumber, sortOrder, true);
    }

    public void Add(String elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks)
    {
        this.Add(elementNumber, sortOrder, ignoreBlanks, false);
    }

    public void Add(
        String elementNumber,
        XLSortOrder sortOrder,
        Boolean ignoreBlanks,
        Boolean matchCase
    )
    {
        this.elements.Add(
            new XLSortElement(
                XlsxSharp.XLHelper.GetColumnNumberFromLetter(elementNumber),
                sortOrder,
                ignoreBlanks,
                matchCase
            )
        );
    }

    public IEnumerator<IXLSortElement> GetEnumerator()
    {
        return this.elements.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public void Clear()
    {
        this.elements.Clear();
    }

    public void Remove(Int32 elementNumber)
    {
        this.elements.RemoveAt(elementNumber - 1);
    }

    internal void AddRange(IEnumerable<XLSortElement> sortElements) =>
        this.elements.AddRange(sortElements);
}
