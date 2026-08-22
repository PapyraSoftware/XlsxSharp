using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.PivotStyleFormats;

/// <summary>
/// A base class for pivot styling API. It has takes a selected <see cref="XLPivotArea"/>
/// and applies the style using <c>.Style*</c> API. The derived classes are responsible for
/// exposing API so user can define an area and then create the desired area (from what user
/// specified) through <see cref="GetCurrentArea"/> method.
/// </summary>
internal abstract class XLPivotStyleFormatBase : IXLPivotStyleFormat
{
    protected readonly XLPivotTable PivotTable;

    protected XLPivotStyleFormatBase(XLPivotTable pivotTable)
    {
        this.PivotTable = pivotTable;
    }

    #region IXLPivotStyleFormat members

    public XLPivotStyleFormatElement AppliesTo { get; init; } = XLPivotStyleFormatElement.Data;

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    #endregion IXLPivotStyleFormat members

    // TODO Styles: Ensure that each pivot area is there only once in a pivot table. Ensure it on load and during modifications.
    internal XLDxFormat Format =>
        new(this.PivotTable.Worksheet.Workbook.Styles, this.GetFormats().First());

    internal abstract XLPivotArea GetCurrentArea();

    internal abstract bool Filter(XLPivotArea area);

    private IEnumerable<XLPivotFormat> GetFormats()
    {
        bool exists = false;
        foreach (XLPivotFormat format in this.PivotTable.Formats)
        {
            if (format.Action == XLPivotFormatAction.Formatting && this.Filter(format.PivotArea))
            {
                exists = true;
                yield return format;
            }
        }

        if (!exists)
        {
            XLPivotFormat format = new(this.GetCurrentArea()) { FormatValue = null };
            this.PivotTable.AddFormat(format);
            yield return format;
        }
    }
}
