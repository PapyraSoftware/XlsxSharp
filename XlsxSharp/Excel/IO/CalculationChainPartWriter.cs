#nullable disable

using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.IO.OpenXmlConst;
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
        // The part is created before the chain is known to be non-empty, and dropped again
        // below when it turns out to be. That order matters: creating it takes the next
        // relationship id, and the ids of every part created after this one depend on whether
        // this one took one.
        if (workbookPart.CalculationChainPart is null)
        {
            workbookPart.AddNewPart<CalculationChainPart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );
        }

        List<(string CellReference, int SheetId, bool IsArrayHead)> chain = BuildChain(workbook);

        // Excel does not keep an empty calcChain part around, and neither should we.
        if (chain.Count == 0)
        {
            workbookPart.DeletePart(workbookPart.CalculationChainPart);
            return;
        }

        XmlWriterSettings settings = new() { Encoding = XlsxSharp.XLHelper.NoBomUTF8 };

        using Stream partStream = workbookPart.CalculationChainPart.GetStream(FileMode.Create);
        using XmlWriter xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartDocument();

        // The "x" prefix rather than a default namespace, which is what Excel and the
        // rest of our writers emit for this part.
        xml.WriteStartElement("x", "calcChain", Main2006SsNs);

        foreach ((string cellReference, int sheetId, bool isArrayHead) in chain)
        {
            xml.WriteStartElement("x", "c", Main2006SsNs);
            xml.WriteAttributeString("r", cellReference);
            xml.WriteAttributeString("i", sheetId.ToInvariantString());
            if (isArrayHead)
            {
                xml.WriteAttributeString("a", TrueValue);
            }

            xml.WriteEndElement(); // c
        }

        xml.WriteEndElement(); // calcChain
        xml.WriteEndDocument();
    }

    /// <summary>
    /// The cells that go into the calculation chain, in the order Excel expects them.
    /// </summary>
    private static List<(string CellReference, int SheetId, bool IsArrayHead)> BuildChain(
        XLWorkbook workbook
    )
    {
        List<(string, int, bool)> chain = [];

        foreach (XLWorksheet worksheet in workbook.WorksheetsInternal)
        {
            int sheetId = (int)worksheet.SheetId;
            foreach (
                XLCell c in worksheet.Internals.CellsCollection.GetCells().Where(c => c.HasFormula)
            )
            {
                if (c.Formula.Type == FormulaType.DataTable)
                {
                    // Excel does not put data tables into the calculation chain.
                    continue;
                }

                if (!c.HasArrayFormula)
                {
                    chain.Add((c.Address.ToString(), sheetId, false));
                    continue;
                }

                c.FormulaReference ??= c.AsRange().RangeAddress;

                // Only the cell the array formula is anchored in carries the whole range, so the
                // other cells of the range must not add it a second time.
                if (!c.FormulaReference.FirstAddress.Equals(c.Address))
                {
                    continue;
                }

                chain.Add((c.Address.ToString(), sheetId, true));
                foreach (IXLCell childCell in worksheet.Range(c.FormulaReference).Cells())
                {
                    chain.Add((childCell.Address.ToString(), sheetId, false));
                }
            }
        }

        return chain;
    }
}
