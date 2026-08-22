using ClosedXML.Parser;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.CalcEngine.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is inserted and cells are shifted down/right.
/// </summary>
internal class ReferenceShiftOnInsertRefModVisitor : CopyVisitor
{
    private readonly SheetArea _insertedBookArea;
    private readonly bool _shiftDown;

    internal ReferenceShiftOnInsertRefModVisitor(SheetArea insertedBookArea, bool shiftDown)
    {
        this._insertedBookArea = insertedBookArea;
        this._shiftDown = shiftDown;
    }

    public override TransformedSymbol SheetReference(
        ModContext ctx,
        SymbolRange range,
        string sheet,
        ReferenceArea reference
    )
    {
        return this.ShiftFormulaReferences(ctx, range, sheet, reference);
    }

    public override TransformedSymbol Reference(
        ModContext ctx,
        SymbolRange range,
        ReferenceArea reference
    )
    {
        return this.ShiftFormulaReferences(ctx, range, null, reference);
    }

    private TransformedSymbol ShiftFormulaReferences(
        ModContext ctx,
        SymbolRange range,
        string? referenceSheetName,
        ReferenceArea referenceToShift
    )
    {
        if (
            !XlsxSharp.XLHelper.SheetComparer.Equals(
                this._insertedBookArea.Name,
                referenceSheetName ?? ctx.Sheet
            )
        )
        {
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);
        }

        bool wouldSplitArea = this._shiftDown
            ? !referenceToShift.TryInsertAndShiftDown(
                this._insertedBookArea.Area,
                out ReferenceArea? shiftedReference
            )
            : !referenceToShift.TryInsertAndShiftRight(
                this._insertedBookArea.Area,
                out shiftedReference
            );

        // Return original reference if the shift would cause a split
        if (wouldSplitArea)
        {
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);
        }

        // If reference was shifted out of sheet, return #REF!
        if (shiftedReference is null)
        {
            return TransformedSymbol.ToText(ctx.Formula, range, XlsxSharp.XLHelper.RefError);
        }

        // Do not allocate a new string unless necessary
        if (referenceToShift == shiftedReference.Value)
        {
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);
        }

        string shiftedReferenceA1 = shiftedReference.Value.GetDisplayStringA1(referenceSheetName);
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReferenceA1);
    }
}
