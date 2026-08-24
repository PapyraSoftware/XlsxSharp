#nullable disable

namespace XlsxSharp.Excel.Sort;

internal class XLSortElements : IXLSortElements
{
    private List<IXLSortElement> elements = [];

    public void Add(int elementNumber) => this.Add(elementNumber, XLSortOrder.Ascending);

    public void Add(int elementNumber, XLSortOrder sortOrder) =>
        this.Add(elementNumber, sortOrder, true);

    public void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks) =>
        this.Add(elementNumber, sortOrder, ignoreBlanks, false);

    public void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase) =>
        this.elements.Add(new XLSortElement(elementNumber, sortOrder, ignoreBlanks, matchCase));

    public void Add(string elementNumber) => this.Add(elementNumber, XLSortOrder.Ascending);

    public void Add(string elementNumber, XLSortOrder sortOrder) =>
        this.Add(elementNumber, sortOrder, true);

    public void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks) =>
        this.Add(elementNumber, sortOrder, ignoreBlanks, false);

    public void Add(
        string elementNumber,
        XLSortOrder sortOrder,
        bool ignoreBlanks,
        bool matchCase
    ) =>
        this.elements.Add(
            new XLSortElement(
                XlsxSharp.XLHelper.GetColumnNumberFromLetter(elementNumber),
                sortOrder,
                ignoreBlanks,
                matchCase
            )
        );

    public IEnumerator<IXLSortElement> GetEnumerator() => this.elements.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    public void Clear() => this.elements.Clear();

    public void Remove(int elementNumber) => this.elements.RemoveAt(elementNumber - 1);

    internal void AddRange(IEnumerable<XLSortElement> sortElements) =>
        this.elements.AddRange(sortElements);
}
