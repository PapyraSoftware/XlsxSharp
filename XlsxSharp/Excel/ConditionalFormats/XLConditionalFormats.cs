using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Parser;

namespace XlsxSharp.Excel.ConditionalFormats;

/// <summary>
/// A container for conditional formatting of a <see cref="XLWorksheet"/>. It contains
/// a collection of <see cref="XLConditionalFormat"/>. Doesn't contain pivot table formats,
/// they are in pivot table <see cref="XLPivotTable.ConditionalFormats"/>,
/// </summary>
internal class XLConditionalFormats
    : IXLConditionalFormats,
        IEnumerable<XLConditionalFormat>,
        ISheetListener
{
    private readonly XLWorksheet _worksheet;
    private readonly List<XLConditionalFormat> _conditionalFormats = [];

    private static readonly List<XLConditionalFormatType> CFTypesExcludedFromConsolidation =
    [
        XLConditionalFormatType.DataBar,
        XLConditionalFormatType.ColorScale,
        XLConditionalFormatType.IconSet,
        XLConditionalFormatType.Top10,
        XLConditionalFormatType.AboveAverage,
        XLConditionalFormatType.IsDuplicate,
        XLConditionalFormatType.IsUnique,
    ];

    public XLConditionalFormats(XLWorksheet worksheet) => this._worksheet = worksheet;

    public void Add(IXLConditionalFormat conditionalFormat)
    {
        XLConditionalFormat addedCf = (XLConditionalFormat)conditionalFormat;
        this._conditionalFormats.Add(addedCf);
        if (addedCf.FormatValue is { } dxf)
        {
            this._worksheet.Workbook.Styles.RegisterDxFormat(dxf);
        }
    }

    public IEnumerator<XLConditionalFormat> GetEnumerator() =>
        this._conditionalFormats.GetEnumerator();

    IEnumerator<IXLConditionalFormat> IEnumerable<IXLConditionalFormat>.GetEnumerator() =>
        this.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    public void Remove(Predicate<IXLConditionalFormat> predicate) =>
        this._conditionalFormats.RemoveAll(predicate);

    public void RemoveAll() => this._conditionalFormats.Clear();

    /// <summary>
    /// Reorders the according to original priority. Done during load process
    /// </summary>
    public void ReorderAccordingToOriginalPriority()
    {
        List<XLConditionalFormat> reorderedFormats =
        [
            .. this._conditionalFormats.OrderBy(cf => cf.Priority),
        ];
        this._conditionalFormats.Clear();
        this._conditionalFormats.AddRange(reorderedFormats);
    }

    /// <summary>
    /// Clear conditional formats in the <paramref name="area"/>. Split if necessary, remove if
    /// conditional format has no area left.
    /// </summary>
    internal void Clear(Area area)
    {
        for (int i = this._conditionalFormats.Count - 1; i >= 0; --i)
        {
            XLConditionalFormat conditionalFormat = this._conditionalFormats[i];
            if (!conditionalFormat.Areas.IntersectsWith(area))
            {
                continue;
            }

            XLAreaList remainingAreas = conditionalFormat.Areas.Excluding(area);
            if (remainingAreas.Count > 0)
            {
                conditionalFormat.Areas = remainingAreas;
            }
            else
            {
                this._conditionalFormats.RemoveAt(i);
            }
        }
    }

    internal void CopyFrom(
        XLWorksheet sourceSheet,
        Area sourceArea,
        Point targetPoint,
        bool mergeUncoveredInSameSheet = false
    )
    {
        // If source and target sheets are same, do not go over the end
        int sourceCfCount = sourceSheet.ConditionalFormats._conditionalFormats.Count;
        for (int i = 0; i < sourceCfCount; ++i)
        {
            XLConditionalFormat sourceCf = sourceSheet.ConditionalFormats._conditionalFormats[i];
            if (!sourceCf.Areas.TryCopyAreaTo(targetPoint, sourceArea, out XLAreaList? targetAreas))
            {
                continue;
            }

            // Legacy behavior where a copied single point was merged into CF when not covered.
            // But only for cell copy API, nor range copy API (even if range is only 1x1).
            if (mergeUncoveredInSameSheet && this._worksheet == sourceSheet)
            {
                foreach (Area targetArea in targetAreas)
                {
                    bool isCovered = sourceCf.Areas.Any(sourceCfArea =>
                        sourceCfArea.Covers(targetArea)
                    );
                    if (!isCovered)
                    {
                        sourceCf.Areas = sourceCf.Areas.With(targetArea);
                    }
                }
            }
            else
            {
                XLConditionalFormat targetCfCopy = new(this._worksheet, sourceCf, targetAreas);
                this.Add(targetCfCopy);
            }
        }
    }

    /// <summary>
    /// The method consolidate the same conditional formats, which are located in adjacent ranges.
    /// </summary>
    internal void Consolidate()
    {
        List<XLConditionalFormat> formats =
        [
            .. this._conditionalFormats.Where(cf => cf.Ranges.Any()),
        ];
        this._conditionalFormats.Clear();

        while (formats.Count > 0)
        {
            XLConditionalFormat format = formats[0];
            if (!CFTypesExcludedFromConsolidation.Contains(format.ConditionalFormatType))
            {
                (List<int> rulesToConsolidate, XLAreaList areasWithSameFormat) =
                    GetConsolidatableRules(formats);
                XLAreaList consolidatedAreas = areasWithSameFormat.GetConsolidated();
                XLConditionalFormat consolidatedCf = new(
                    this._worksheet,
                    format,
                    consolidatedAreas
                );

                // Remove consolidated formats
                rulesToConsolidate.Reverse();
                foreach (int consolidatedRuleIndex in rulesToConsolidate)
                {
                    formats.RemoveAt(consolidatedRuleIndex);
                }

                format = consolidatedCf;
            }

            this._conditionalFormats.Add(format);
            formats.RemoveAt(0);
        }
    }

    private static (List<int> RulesToConsolidate, XLAreaList AreaList) GetConsolidatableRules(
        List<XLConditionalFormat> conditionalFormats
    )
    {
        XLConditionalFormat rule = conditionalFormats[0];
        List<Area> sameFormatAreas = [.. rule.Areas];
        List<Area> differentFormatAreas = [];

        // The ids to the list must be in the ascending order
        List<int> rulesToConsolidate = [];
        for (int i = 1; i < conditionalFormats.Count; ++i)
        {
            XLConditionalFormat candidateRule = conditionalFormats[i];

            bool intersectsDifferentFormatAreas = differentFormatAreas.Any(differentFormatArea =>
                candidateRule.Areas.Any(v => v.Intersects(differentFormatArea))
            );
            if (intersectsDifferentFormatAreas)
            {
                // We reached a rule intersecting any of captured ranges. Stop for not breaking the priorities.
                break;
            }

            bool isSameFormat = XLConditionalFormat.NoRangeComparer.Equals(candidateRule, rule);
            if (isSameFormat)
            {
                // We reached a rule that has same format as the consolidated rule and doesn't intersect different
                // format areas. We can consolidate the candidate rule with the rule without potentially breaking
                // any rule with a priority between rule and candidate rule.
                sameFormatAreas.AddRange(candidateRule.Areas);
                rulesToConsolidate.Add(i);
                continue;
            }

            bool intersectsSameFormatAreas = sameFormatAreas.Any(sameFormatArea =>
                candidateRule.Areas.Any(v => v.Intersects(sameFormatArea))
            );
            if (intersectsSameFormatAreas)
            {
                // We reached a rule that has different format and intersects area to be consolidated. That means
                // it's not possible to consolidate any subsequent rule, because it could break this one, and
                // consolidation must stop here.
                break;
            }

            // The most common case: The candidate rule has a different format and doesn't intersect the sameFormatAreas
            // The format thus must be added to the differentFormatAreas, because it can interrupt subsequent rules.
            differentFormatAreas.AddRange(candidateRule.Areas);
        }

        return (rulesToConsolidate, new XLAreaList(sameFormatAreas));
    }

    #region ISheetListener

    void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, Area insertedArea)
    {
        SheetArea inserted = new(sheet.Name, insertedArea);
        ReferenceShiftOnInsertRefModVisitor refMod = new(inserted, true);
        this.AdjustFormulas(refMod);

        this.AdjustConditionalFormatAreas(
            sheet,
            inserted.Area,
            static (sqref, insertedArea) => sqref.InsertAndShiftDown(insertedArea)
        );
    }

    void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, Area insertedArea)
    {
        SheetArea inserted = new(sheet.Name, insertedArea);
        ReferenceShiftOnInsertRefModVisitor refMod = new(inserted, false);
        this.AdjustFormulas(refMod);

        this.AdjustConditionalFormatAreas(
            sheet,
            inserted.Area,
            static (sqref, insertedArea) => sqref.InsertAndShiftRight(insertedArea)
        );
    }

    void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea)
    {
        SheetArea deleted = new(sheet.Name, deletedArea);
        ReferenceShiftOnDeleteRefModVisitor refMod = new(
            deleted,
            XLShiftDeletedCells.ShiftCellsLeft
        );
        this.AdjustFormulas(refMod);

        this.AdjustConditionalFormatAreas(
            sheet,
            deleted.Area,
            static (sqref, deletedArea) => sqref.DeleteAndShiftLeft(deletedArea)
        );
    }

    void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea)
    {
        SheetArea deleted = new(sheet.Name, deletedArea);
        ReferenceShiftOnDeleteRefModVisitor refMod = new(deleted, XLShiftDeletedCells.ShiftCellsUp);
        this.AdjustFormulas(refMod);

        this.AdjustConditionalFormatAreas(
            sheet,
            deleted.Area,
            static (sqref, deletedArea) => sqref.DeleteAndShiftUp(deletedArea)
        );
    }

    private void AdjustFormulas(CopyVisitor refMod)
    {
        foreach (XLConditionalFormat conditionalFormat in this._conditionalFormats)
        {
            Point anchor = conditionalFormat.Areas[0].FirstPoint;
            int[] formulaIndexes =
            [
                .. conditionalFormat.Values.Where(x => x.Value.IsFormula).Select(x => x.Key),
            ];
            foreach (int index in formulaIndexes)
            {
                XLFormula originalFormula = conditionalFormat.Values[index];
                string shiftedFormula = FormulaConverter.ModifyA1(
                    originalFormula.Value,
                    this._worksheet.Name,
                    anchor.Row,
                    anchor.Column,
                    refMod
                );
                conditionalFormat.Values[index] = new XLFormula(shiftedFormula)
                {
                    IsFormula = true,
                };
            }
        }
    }

    private void AdjustConditionalFormatAreas(
        XLWorksheet sheet,
        Area affectedRange,
        Func<XLAreaList, Area, XLAreaList> adjustAreas
    )
    {
        if (sheet != this._worksheet)
        {
            return;
        }

        for (int i = this._conditionalFormats.Count - 1; i >= 0; --i)
        {
            XLConditionalFormat conditionalFormat = this._conditionalFormats[i];
            XLAreaList modifiedAreaList = adjustAreas(conditionalFormat.Areas, affectedRange);
            if (modifiedAreaList.Count == 0)
            {
                this._conditionalFormats.RemoveAt(i);
            }
            else
            {
                conditionalFormat.Areas = modifiedAreaList;
            }
        }
    }

    #endregion
}
