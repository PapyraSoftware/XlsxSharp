using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Parser;
using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

/// <summary>
/// A collection of a named ranges, either for workbook or for worksheet.
/// </summary>
internal class XLDefinedNames : IXLDefinedNames, IEnumerable<XLDefinedName>, ISheetListener
{
    private readonly Dictionary<String, XLDefinedName> _namedRanges = new(
        XlsxSharp.XLHelper.NameComparer
    );

    internal XLWorkbook Workbook { get; }

    internal XLWorksheet? Worksheet { get; }

    internal XLNamedRangeScope Scope { get; }

    public XLDefinedNames(XLWorksheet worksheet)
        : this(worksheet.Workbook)
    {
        this.Worksheet = worksheet;
        this.Scope = XLNamedRangeScope.Worksheet;
    }

    public XLDefinedNames(XLWorkbook workbook)
    {
        this.Workbook = workbook;
        this.Scope = XLNamedRangeScope.Workbook;
    }

    #region IXLNamedRanges Members

    [Obsolete]
    IXLDefinedName IXLDefinedNames.NamedRange(String name) => this.DefinedName(name);

    IXLDefinedName IXLDefinedNames.DefinedName(String name) => this.DefinedName(name);

    internal XLDefinedName DefinedName(String name)
    {
        if (this._namedRanges.TryGetValue(name, out XLDefinedName range))
        {
            return range;
        }

        throw new KeyNotFoundException($"Name {name} not found.");
    }

    public IXLDefinedName Add(String name, String rangeAddress) =>
        this.Add(name, rangeAddress, null);

    public IXLDefinedName Add(String name, IXLRange range) => this.Add(name, range, null);

    public IXLDefinedName Add(String name, IXLRanges ranges) => this.Add(name, ranges, null);

    public IXLDefinedName Add(String name, String rangeAddress, String? comment) =>
        this.Add(name, rangeAddress, comment, validateName: true, validateRangeAddress: true);

    /// <summary>
    /// Adds the specified range name.
    /// </summary>
    /// <param name="name">Name of the range.</param>
    /// <param name="rangeAddress">The range address.</param>
    /// <param name="comment">The comment.</param>
    /// <param name="validateName">if set to <c>true</c> validates the name.</param>
    /// <param name="validateRangeAddress">if set to <c>true</c> range address will be checked for validity.</param>
    /// <exception cref="NotSupportedException"></exception>
    /// <exception cref="ArgumentException">
    /// For named ranges in the workbook scope, specify the sheet name in the reference.
    /// </exception>
    internal IXLDefinedName Add(
        String name,
        String rangeAddress,
        String? comment,
        Boolean validateName,
        Boolean validateRangeAddress
    )
    {
        // When loading named ranges from an existing file, we do not validate the range address or name.
        if (validateRangeAddress)
        {
            Match match = XlsxSharp.XLHelper.NamedRangeReferenceRegex.Match(rangeAddress);

            if (!match.Success)
            {
                if (XlsxSharp.XLHelper.IsValidRangeAddress(rangeAddress))
                {
                    IXLRange? range;
                    if (this.Scope == XLNamedRangeScope.Worksheet)
                    {
                        range = this.Worksheet!.Range(rangeAddress);
                    }
                    else if (this.Scope == XLNamedRangeScope.Workbook)
                    {
                        range = this.Workbook.Range(rangeAddress);
                    }
                    else
                    {
                        throw new NotSupportedException($"Scope {this.Scope} is not supported");
                    }

                    if (range == null)
                    {
                        throw new ArgumentException(
                            string.Format(
                                "The range address '{0}' for the named range '{1}' is not a valid range.",
                                rangeAddress,
                                name
                            )
                        );
                    }

                    if (
                        this.Scope == XLNamedRangeScope.Workbook
                        || !XlsxSharp
                            .XLHelper.NamedRangeReferenceRegex.Match(range.ToString())
                            .Success
                    )
                    {
                        throw new ArgumentException(
                            "For named ranges in the workbook scope, specify the sheet name in the reference."
                        );
                    }

                    rangeAddress = range.ToString();
                }
            }
        }

        XLDefinedName namedRange = new(this, name, validateName, rangeAddress, comment);
        this._namedRanges.Add(name, namedRange);
        return namedRange;
    }

    public IXLDefinedName Add(String name, IXLRange range, String? comment)
    {
        XLRanges ranges = new(this.Workbook) { range };
        return this.Add(name, ranges, comment);
    }

    public IXLDefinedName Add(String name, IXLRanges ranges, String? comment)
    {
        string formula = string.Join(
            ",",
            ranges.Select(r => r.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true))
        );
        XLDefinedName namedRange = new(this, name, true, formula, comment);
        this._namedRanges.Add(name, namedRange);
        return namedRange;
    }

    internal XLDefinedName Add(String name, XLDefinedName namedRange)
    {
        this._namedRanges.Add(name, namedRange);
        return namedRange;
    }

    public void Delete(String rangeName) => this._namedRanges.Remove(rangeName);

    public void Delete(Int32 rangeIndex) =>
        this._namedRanges.Remove(this._namedRanges.ElementAt(rangeIndex).Key);

    public void DeleteAll() => this._namedRanges.Clear();

    /// <summary>
    /// Returns a subset of named ranges that do not have invalid references.
    /// </summary>
    public IEnumerable<IXLDefinedName> ValidNamedRanges() =>
        this._namedRanges.Values.Where(nr => nr.IsValid);

    /// <summary>
    /// Returns a subset of named ranges that do have invalid references.
    /// </summary>
    public IEnumerable<IXLDefinedName> InvalidNamedRanges() =>
        this._namedRanges.Values.Where(nr => !nr.IsValid);

    #endregion IXLNamedRanges Members

    IEnumerator<XLDefinedName> IEnumerable<XLDefinedName>.GetEnumerator() => this.GetEnumerator();

    IEnumerator<IXLDefinedName> IEnumerable<IXLDefinedName>.GetEnumerator() => this.GetEnumerator();

    public Dictionary<string, XLDefinedName>.ValueCollection.Enumerator GetEnumerator() =>
        this._namedRanges.Values.GetEnumerator();

    #region IEnumerable Members

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    #endregion IEnumerable Members

    public Boolean TryGetValue(String name, [NotNullWhen(true)] out IXLDefinedName? definedName)
    {
        if (this.TryGetScopedValue(name, out XLDefinedName? sheetDefinedName))
        {
            definedName = sheetDefinedName;
            return true;
        }

        definedName =
            this.Scope == XLNamedRangeScope.Workbook ? this.Workbook.DefinedName(name) : null;

        return definedName is not null;
    }

    internal Boolean TryGetScopedValue(
        String name,
        [NotNullWhen(true)] out XLDefinedName? definedName
    )
    {
        if (this._namedRanges.TryGetValue(name, out definedName))
        {
            return true;
        }

        return false;
    }

    public Boolean Contains(String name)
    {
        if (this._namedRanges.ContainsKey(name))
        {
            return true;
        }

        if (this.Scope == XLNamedRangeScope.Workbook)
        {
            return this.Workbook.DefinedName(name) is not null;
        }
        else
        {
            return false;
        }
    }

    internal void OnWorksheetDeleted(string worksheetName) =>
        this._namedRanges.Values.ForEach(nr => nr.OnWorksheetDeleted(worksheetName));

    #region ISheetListner

    void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, Area insertedArea)
    {
        SheetArea insertedBookArea = new(sheet.Name, insertedArea);
        ReferenceShiftOnInsertRefModVisitor refMod = new(insertedBookArea, true);
        this.ShiftReferences(refMod);
    }

    void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, Area insertedArea)
    {
        SheetArea insertedBookArea = new(sheet.Name, insertedArea);
        ReferenceShiftOnInsertRefModVisitor refMod = new(insertedBookArea, false);
        this.ShiftReferences(refMod);
    }

    void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea)
    {
        SheetArea deletedBookArea = new(sheet.Name, deletedArea);
        ReferenceShiftOnDeleteRefModVisitor refMod = new(
            deletedBookArea,
            XLShiftDeletedCells.ShiftCellsLeft
        );
        this.ShiftReferences(refMod);
    }

    void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea)
    {
        SheetArea deletedBookArea = new(sheet.Name, deletedArea);
        ReferenceShiftOnDeleteRefModVisitor refMod = new(
            deletedBookArea,
            XLShiftDeletedCells.ShiftCellsUp
        );
        this.ShiftReferences(refMod);
    }

    private void ShiftReferences(CopyVisitor refMod)
    {
        foreach (XLDefinedName definedName in this._namedRanges.Values)
        {
            string nameFormula = definedName.RefersTo;

            // Defined name formula should never rely on a context info. Formula should contain
            // only absolute references with a sheet -> use empty sheet that should never match
            // and the cell point is thus never used and can be left at A1.
            string shiftedFormula = FormulaConverter.ModifyA1(
                nameFormula,
                string.Empty,
                1,
                1,
                refMod
            );
            definedName.SetRefersTo(shiftedFormula);
        }
    }
    #endregion
}
