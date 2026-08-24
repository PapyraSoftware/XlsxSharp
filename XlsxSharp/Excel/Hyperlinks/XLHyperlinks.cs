using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel;

internal class XLHyperlinks : IXLHyperlinks, ISheetListener
{
    private readonly XLWorksheet _worksheet;

    /// <summary>
    /// XLHyperlink doesn't contain range, it is user created and only then it is associated with an area in a sheet.
    /// </summary>
    private readonly List<(XLHyperlink Link, Area Area)> _hyperlinks = [];
    private readonly RTree<XLHyperlink> _areaIndex = new();
    private readonly Dictionary<XLHyperlink, Area> _linkIndex = new();

    private delegate (bool Success, Area? RepositionedArea) RepositionFunc(Area hyperlinkArea);

    internal XLHyperlinks(XLWorksheet worksheet) => this._worksheet = worksheet;

    internal string WorksheetName => this._worksheet.Name;

    #region ISheetListener

    void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, Area insertedArea) =>
        this.RepositionOnChange(
            sheet,
            hyperlinkArea =>
            {
                bool success = hyperlinkArea.TryInsertAreaAndShiftDown(
                    insertedArea,
                    out Area? newHlArea
                );
                return (success, newHlArea);
            }
        );

    void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, Area insertedArea) =>
        this.RepositionOnChange(
            sheet,
            hyperlinkArea =>
            {
                bool success = hyperlinkArea.TryInsertAreaAndShiftRight(
                    insertedArea,
                    out Area? newHlArea
                );
                return (success, newHlArea);
            }
        );

    void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea) =>
        this.RepositionOnChange(
            sheet,
            hyperlinkArea =>
            {
                bool success = hyperlinkArea.TryDeleteAreaAndShiftLeft(
                    deletedArea,
                    out Area? newHlArea
                );
                return (success, newHlArea);
            }
        );

    void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea) =>
        this.RepositionOnChange(
            sheet,
            hyperlinkArea =>
            {
                bool success = hyperlinkArea.TryDeleteAreaAndShiftUp(
                    deletedArea,
                    out Area? newHlArea
                );
                return (success, newHlArea);
            }
        );

    private void RepositionOnChange(XLWorksheet sheet, RepositionFunc reposition)
    {
        if (sheet != this._worksheet)
        {
            return;
        }

        // Styles are responsibility of style slice => only shift areas
        foreach ((XLHyperlink link, Area linkArea) in this._hyperlinks.ToArray())
        {
            (bool success, Area? newLinkArea) = reposition(linkArea);
            if (!success)
            {
                continue; // Partial cover, don't move.
            }

            if (linkArea == newLinkArea)
            {
                continue; // Nothing changed
            }

            this.Remove(link);
            if (newLinkArea is not null)
            {
                this.Add(newLinkArea.Value, link);
            }
        }
    }

    #endregion ISheetListener

    public IEnumerator<XLHyperlink> GetEnumerator() =>
        // Enumerate in same order it was loaded and will be saved
        this._hyperlinks.Select(static x => x.Link).GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    /// <inheritdoc />
    public bool Delete(XLHyperlink hyperlink)
    {
        if (!this.Remove(hyperlink, out Area linkArea))
        {
            return false;
        }

        this.ClearHyperlinkStyle(linkArea);
        return true;
    }

    /// <inheritdoc />
    public bool Delete(IXLAddress address)
    {
        if (address.Worksheet is not null && address.Worksheet != this._worksheet)
        {
            return false;
        }

        Point cellPoint = Point.FromAddress(address);
        if (!this.TryGet(cellPoint, out XLHyperlink? cellLink))
        {
            return false;
        }

        this.Remove(cellLink);
        this.ClearHyperlinkStyle(cellPoint);
        return true;
    }

    /// <inheritdoc />
    public XLHyperlink Get(IXLAddress address)
    {
        if (address.Worksheet is not null && address.Worksheet != this._worksheet)
        {
            throw new KeyNotFoundException("Address is for a different sheet.");
        }

        Point point = Point.FromAddress(address);
        if (!this.TryGet(point, out XLHyperlink? link))
        {
            throw new KeyNotFoundException($"No hyperlink is defined for cell {point}.");
        }

        return link;
    }

    /// <inheritdoc />
    public bool TryGet(IXLAddress address, [NotNullWhen(true)] out XLHyperlink? hyperlink)
    {
        if (address.Worksheet is not null && address.Worksheet != this._worksheet)
        {
            hyperlink = null;
            return false;
        }

        Point point = Point.FromAddress(address);
        return this.TryGet(point, out hyperlink);
    }

    internal bool HasHyperlink(Point point)
    {
        List<RTree<XLHyperlink>.Node> areaNodes = [];
        return this._areaIndex.GetNodes(point, areaNodes).Count > 0;
    }

    /// <summary>
    /// Set a hyperlink of a single cell. Doesn't modify style, ignores hyperlinks with areas that
    /// cover the cell.
    /// </summary>
    internal void SetCellHyperlink(Point point, XLHyperlink? link)
    {
        // We only care about links defined for individual cell, not any link that covers the cell
        List<RTree<XLHyperlink>.Node> pointNodes = [];
        this._areaIndex.GetNodes(point, pointNodes);
        foreach (RTree<XLHyperlink>.Node existingLink in pointNodes)
        {
            this.Remove(existingLink.Data);
        }

        if (link is null)
        {
            return;
        }

        this.Add(point, link);
    }

    internal bool TryGet(Point point, [NotNullWhen(true)] out XLHyperlink? hyperlink)
    {
        Area cellArea = new(point);
        List<RTree<XLHyperlink>.Node> areaNodes = [];
        this._areaIndex.GetNodes(cellArea, areaNodes);

        if (areaNodes.Count == 0)
        {
            hyperlink = null;
            return false;
        }

        if (areaNodes.Count == 1)
        {
            hyperlink = areaNodes[0].Data;
            return true;
        }

        // There are multiple areas for the point. When hyperlink areas overlap, Excel opens
        // the last one. So it is likely the correct one. But take a random one (areaNodes are
        // not guaranteed to be in correct order), because this API is just beyond any hope and
        // will be completely scrapped ASAP.
        hyperlink = areaNodes[^1].Data;
        return true;
    }

    internal XLCell? GetCell(XLHyperlink hyperlink)
    {
        if (!this._linkIndex.TryGetValue(hyperlink, out Area area))
        {
            return null;
        }

        return new XLCell(this._worksheet, area.FirstPoint);
    }

    private void Add(Area linkArea, XLHyperlink link)
    {
        if (link.Container is not null && link.Container != this)
        {
            throw new InvalidOperationException(
                "Hyperlink is attached to a different worksheet. Either remove it from the original worksheet or create a new hyperlink."
            );
        }

        if (this._linkIndex.ContainsKey(link))
        {
            return;
        }

        this._linkIndex.Add(link, linkArea);
        this._areaIndex.Insert(new RTree<XLHyperlink>.Node(linkArea, link));
        this._hyperlinks.Add((link, linkArea));
        link.Container = this;
        Debug.Assert(this._hyperlinks.Count == this._linkIndex.Count);
        Debug.Assert(this._hyperlinks.Count == this._areaIndex.Count);
    }

    private void Remove(XLHyperlink link) => this.Remove(link, out _);

    private bool Remove(XLHyperlink link, out Area area)
    {
        if (!this._linkIndex.Remove(link, out area))
        {
            return false;
        }

        this._areaIndex.Delete(new RTree<XLHyperlink>.Node(area, link));
        this._hyperlinks.RemoveAll(x => x.Link == link);
        link.Container = null;
        Debug.Assert(this._hyperlinks.Count == this._linkIndex.Count);
        Debug.Assert(this._hyperlinks.Count == this._areaIndex.Count);
        return true;
    }

    private void ClearHyperlinkStyle(Area range)
    {
        XLFontFormatValue worksheetFont = this._worksheet.GetFormat().Font;
        XLColor sheetColor = worksheetFont.Color;
        XLFontUnderlineValues sheetUnderline = worksheetFont.Underline;
        foreach (Point point in range)
        {
            XLCell? cell = this._worksheet.GetCell(point);
            if (cell is null)
            {
                continue;
            }

            if (cell.Style.Font.FontColor.Equals(XLColor.FromTheme(XLThemeColor.Hyperlink)))
            {
                cell.Style.Font.FontColor = sheetColor;
            }

            cell.Style.Font.Underline = sheetUnderline;
        }
    }
}
