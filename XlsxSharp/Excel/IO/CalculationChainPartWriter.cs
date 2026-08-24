#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static XlsxSharp.Excel.XLWorkbook;

namespace XlsxSharp.Excel.IO;

internal class CalculationChainPartWriter
{
    internal static void GenerateContent(
        WorkbookPart workbookPart,
        XLWorkbook workbook,
        SaveContext context
    )
    {
        if (workbookPart.CalculationChainPart == null)
        {
            workbookPart.AddNewPart<CalculationChainPart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );
        }

        if (workbookPart.CalculationChainPart.CalculationChain == null)
        {
            workbookPart.CalculationChainPart.CalculationChain = new CalculationChain();
        }

        CalculationChain calculationChain = workbookPart.CalculationChainPart.CalculationChain;
        calculationChain.RemoveAllChildren<CalculationCell>();

        foreach (XLWorksheet worksheet in workbook.WorksheetsInternal)
        {
            foreach (
                XLCell c in worksheet.Internals.CellsCollection.GetCells().Where(c => c.HasFormula)
            )
            {
                if (c.Formula.Type == FormulaType.DataTable)
                {
                    // Do nothing, Excel doesn't generate calc chain for data table
                }
                else if (c.HasArrayFormula)
                {
                    if (c.FormulaReference == null)
                    {
                        c.FormulaReference = c.AsRange().RangeAddress;
                    }

                    if (c.FormulaReference.FirstAddress.Equals(c.Address))
                    {
                        CalculationCell cc = new()
                        {
                            CellReference = c.Address.ToString(),
                            SheetId = (int)worksheet.SheetId,
                        };

                        cc.Array = true;
                        calculationChain.AppendChild(cc);

                        foreach (IXLCell childCell in worksheet.Range(c.FormulaReference).Cells())
                        {
                            calculationChain.AppendChild(
                                new CalculationCell
                                {
                                    CellReference = childCell.Address.ToString(),
                                    SheetId = (int)worksheet.SheetId,
                                }
                            );
                        }
                    }
                }
                else
                {
                    calculationChain.AppendChild(
                        new CalculationCell
                        {
                            CellReference = c.Address.ToString(),
                            SheetId = (int)worksheet.SheetId,
                        }
                    );
                }
            }
        }

        if (!calculationChain.Any())
        {
            workbookPart.DeletePart(workbookPart.CalculationChainPart);
        }
    }
}
