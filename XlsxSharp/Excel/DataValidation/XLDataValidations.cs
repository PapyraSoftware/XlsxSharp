using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace XlsxSharp.Excel.DataValidation;

internal class XLDataValidations : IXLDataValidations, IEnumerable<XLDataValidation>, ISheetListener
{
    private readonly List<XLDataValidation> _dataValidations = [];
    private readonly XLWorksheet _worksheet;

    public XLDataValidations(XLWorksheet worksheet) =>
        this._worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));

    #region IXLDataValidations Members

    IXLWorksheet IXLDataValidations.Worksheet => this._worksheet;

    IXLDataValidation IXLDataValidations.Add(IXLDataValidation dataValidation)
    {
        ArgumentNullException.ThrowIfNull(dataValidation);

        XLDataValidation dv = (XLDataValidation)dataValidation;
        if (dv.Worksheet != this._worksheet)
        {
            return this.CopyFrom(dv);
        }

        // It's possible that it was detached and while detached, it had added some areas?
        // I have a very hard time understanding the use case and intended behavior. This
        // API should be scrapped.
        if (!this._dataValidations.Contains(dv))
        {
            // Adding a range can split current one -> clear existing DVs so new one can be
            // added and "one DV per cell" is kept.
            foreach (Area area in dv.Areas)
            {
                this.AdjustDataValidationAreas(
                    this._worksheet,
                    area,
                    static (dataValidationAreas, areaOfNewValidation) =>
                        dataValidationAreas.DeleteWithoutShift(areaOfNewValidation)
                );
            }

            this._dataValidations.Add(dv);
        }

        return dv;
    }

    public Boolean ContainsSingle(IXLRange range)
    {
        Int32 count = 0;
        foreach (
            XLDataValidation xlDataValidation in this._dataValidations.Where(dv =>
                dv.Ranges.Contains(range)
            )
        )
        {
            count++;
            if (count > 1)
            {
                return false;
            }
        }

        return count == 1;
    }

    public void Delete(Predicate<IXLDataValidation> predicate)
    {
        List<XLDataValidation> dataValidationsToRemove =
        [
            .. this._dataValidations.Where(dv => predicate(dv)),
        ];

        dataValidationsToRemove.ForEach(this.Delete);
    }

    /// <summary>
    /// Get all data validation rules applied to ranges that intersect the specified range.
    /// </summary>
    public IEnumerable<IXLDataValidation> GetAllInRange(IXLRangeAddress rangeAddress)
    {
        if (rangeAddress is null || !rangeAddress.IsValid)
        {
            yield break;
        }

        if (rangeAddress.Worksheet != this._worksheet)
        {
            yield break;
        }

        Area intersectingArea = Area.FromRangeAddress(rangeAddress);
        foreach (XLDataValidation dataValidation in this._dataValidations)
        {
            foreach (Area area in dataValidation.Areas)
            {
                if (intersectingArea.Intersects(area))
                {
                    yield return dataValidation;
                    break;
                }
            }
        }
    }

    public IEnumerator<XLDataValidation> GetEnumerator() => this._dataValidations.GetEnumerator();

    IEnumerator<IXLDataValidation> IEnumerable<IXLDataValidation>.GetEnumerator() =>
        this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    /// <summary>
    /// Get the data validation rule for the range with the specified address if it exists.
    /// </summary>
    /// <param name="rangeAddress">A range address.</param>
    /// <param name="foundDataValidation">Data validation rule which ranges collection includes the specified
    /// address. The specified range should be fully covered with the data validation rule.
    /// For example, if the rule is applied to ranges A1:A3,C1:C3 then this method will
    /// return True for ranges A1:A3, C1:C2, A2:A3, and False for ranges A1:C3, A1:C1, etc.</param>
    /// <returns>True is the data validation rule was found, false otherwise.</returns>
    public bool TryGet(
        IXLRangeAddress rangeAddress,
        [NotNullWhen(true)] out IXLDataValidation? foundDataValidation
    )
    {
        if (
            rangeAddress is null
            || !rangeAddress.IsValid
            || rangeAddress.Worksheet != this._worksheet
        )
        {
            foundDataValidation = null;
            return false;
        }

        Area coveredArea = Area.FromRangeAddress(rangeAddress);
        foreach (XLDataValidation dataValidation in this._dataValidations)
        {
            foreach (Area area in dataValidation.Areas)
            {
                if (area.Covers(coveredArea))
                {
                    foundDataValidation = dataValidation;
                    return true;
                }
            }
        }

        foundDataValidation = null;
        return false;
    }

    #endregion IXLDataValidations Members

    /// <summary>
    /// Create a new DV with an initial area.
    /// </summary>
    internal XLDataValidation Create(Area area)
    {
        XLDataValidation dv = new(this._worksheet);
        this._dataValidations.Add(dv);
        this.AddArea(dv, area);
        return dv;
    }

    /// <summary>
    /// Create a new DV that is created from another DV from different sheet.
    /// </summary>
    internal XLDataValidation CopyFrom(XLDataValidation original)
    {
        XLDataValidation dv = new(this._worksheet);
        this._dataValidations.Add(dv);
        dv.CopyFrom(original);
        return dv;
    }

    internal void Delete(Area areaToDelete)
    {
        for (int i = this._dataValidations.Count - 1; i >= 0; --i)
        {
            XLDataValidation dataValidation = this._dataValidations[i];
            foreach (Area dataValidationArea in dataValidation.Areas)
            {
                if (dataValidationArea.Intersects(areaToDelete))
                {
                    this._dataValidations.RemoveAt(i);
                    break;
                }
            }
        }
    }

    internal void Delete(XLDataValidation dataValidation) =>
        this._dataValidations.Remove(dataValidation);

    internal void Consolidate()
    {
        Func<IXLDataValidation, IXLDataValidation, bool> areEqual = (dv1, dv2) =>
        {
            return dv1.IgnoreBlanks == dv2.IgnoreBlanks
                && dv1.InCellDropdown == dv2.InCellDropdown
                && dv1.ShowErrorMessage == dv2.ShowErrorMessage
                && dv1.ShowInputMessage == dv2.ShowInputMessage
                && dv1.InputTitle == dv2.InputTitle
                && dv1.InputMessage == dv2.InputMessage
                && dv1.ErrorTitle == dv2.ErrorTitle
                && dv1.ErrorMessage == dv2.ErrorMessage
                && dv1.ErrorStyle == dv2.ErrorStyle
                && dv1.AllowedValues == dv2.AllowedValues
                && dv1.Operator == dv2.Operator
                && dv1.MinValue == dv2.MinValue
                && dv1.MaxValue == dv2.MaxValue
                && dv1.Value == dv2.Value;
        };

        List<XLDataValidation> rules = [.. this._dataValidations];
        this._dataValidations.Clear();

        while (rules.Any())
        {
            XLDataValidation consRule = rules.First();
            this._dataValidations.Add(consRule);
            List<XLDataValidation> similarRules = [.. rules.Where(r => areEqual(consRule, r))];
            similarRules.ForEach(r => rules.Remove(r));

            IXLRanges consolidatedRanges = new XLRanges(this._worksheet);
            foreach (Area similarRuleArea in similarRules.SelectMany(dv => dv.Areas))
            {
                consolidatedRanges.Add(
                    this._worksheet.Range(
                        XLRangeAddress.FromSheetRange(this._worksheet, similarRuleArea)
                    )
                );
            }

            consolidatedRanges = consolidatedRanges.Consolidate();

            consRule.ClearRanges();
            consRule.AddRanges(consolidatedRanges);
        }
    }

    internal void AddArea(XLDataValidation modifiedDataValidation, Area addedArea)
    {
        // Add an area to modifiedDV. This must be done carefully, because there can be only
        // one DV per cell. Due to this problem, the correspondence area-DV should be managed
        // by the DV collection and this method should be private. Change to private would
        // require change of DV to a nested class + separation of API object, so the method is
        // internal + exception.
        if (!this._dataValidations.Contains(modifiedDataValidation))
        {
            throw new ArgumentException(
                "Data validation is not a data validation of this sheet.",
                nameof(modifiedDataValidation)
            );
        }

        // There can be only one DV per cell. Remove DVs from cells that should now belong
        // to the area and remove DVs without any cells.
        for (int i = this._dataValidations.Count - 1; i >= 0; --i)
        {
            XLDataValidation dataValidation = this._dataValidations[i];

            // Area could cover whole modifiedDataValidation and could remove the modifiedDV
            // before the addedArea could be added to the modifiedDV. To avoid this, it is not
            // cleared.
            if (dataValidation == modifiedDataValidation)
            {
                continue;
            }

            dataValidation.Areas = dataValidation.Areas.DeleteWithoutShift(addedArea);
            if (dataValidation.Areas.Count == 0)
            {
                this._dataValidations.RemoveAt(i);
            }
        }

        // Ensure the modifiedDV area list contains only disjunct areas to ensure
        // the "one DV per cell" invariant.
        modifiedDataValidation.Areas = modifiedDataValidation
            .Areas.DeleteWithoutShift(addedArea)
            .With(addedArea);
    }

    void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, Area area) =>
        this.AdjustDataValidationAreas(
            sheet,
            area,
            static (sqref, insertedArea) => sqref.InsertAndShiftDown(insertedArea)
        );

    void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, Area area) =>
        this.AdjustDataValidationAreas(
            sheet,
            area,
            static (sqref, insertedArea) => sqref.InsertAndShiftRight(insertedArea)
        );

    void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedRange) =>
        this.AdjustDataValidationAreas(
            sheet,
            deletedRange,
            static (sqref, deletedArea) => sqref.DeleteAndShiftLeft(deletedArea)
        );

    void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedRange) =>
        this.AdjustDataValidationAreas(
            sheet,
            deletedRange,
            static (sqref, deletedArea) => sqref.DeleteAndShiftUp(deletedArea)
        );

    private void AdjustDataValidationAreas(
        XLWorksheet sheet,
        Area affectedRange,
        Func<XLAreaList, Area, XLAreaList> adjustAreas
    )
    {
        if (sheet != this._worksheet)
        {
            return;
        }

        for (int i = this._dataValidations.Count - 1; i >= 0; --i)
        {
            XLDataValidation dataValidation = this._dataValidations[i];
            XLAreaList modifiedAreaList = adjustAreas(dataValidation.Areas, affectedRange);
            if (modifiedAreaList.Count == 0)
            {
                this._dataValidations.RemoveAt(i);
            }
            else
            {
                dataValidation.Areas = modifiedAreaList;
            }
        }
    }
}
