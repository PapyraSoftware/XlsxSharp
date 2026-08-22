#nullable disable

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

public partial class XLWorkbook
{
    #region Nested type: SaveContext

    internal sealed class SaveContext
    {
        public SaveContext()
        {
            this.RelIdGenerator = new RelIdGenerator();
            this.TableId = 0;
            this.TableNames = [];
            this.PivotSourceCacheId = 0;
        }

        public RelIdGenerator RelIdGenerator { get; }

        /// <summary>
        /// A map of number format to a number format id for saved file. It contains all number
        /// formats from the file, all number formats used in the application (styles, pivot
        /// tables, dxf) and all predefined formats.
        /// </summary>
        internal Dictionary<string, int> NumberFormatMap = new();

        internal Dictionary<XLFontFormatValue, int> FontMap = new();

        internal Dictionary<XLCellFormatValue, uint> FormatMap = new();

        internal Dictionary<XLDxfValue, uint> DxfMap = new();

        public uint TableId { get; set; }
        public HashSet<string> TableNames { get; private set; }

        /// <summary>
        /// A free id that can be used by the workbook to reference to a pivot cache.
        /// The <c>PivotCaches</c> element in a workbook connects the parts with pivot
        /// cache parts.
        /// </summary>
        public uint PivotSourceCacheId { get; set; }

        /// <summary>
        /// A map of shared string ids. The index is the actual index from sharedStringId and
        /// value is an mapped stringId to write to a file. The mapped stringId has no gaps
        /// between ids.
        /// </summary>
        public List<int> SstMap { get; set; }

#nullable enable
        internal int GetSharedStringId(XLCell xlCell, string text)
        {
            int sharedStringId = this.SstMap[xlCell.MemorySstId];
            if (sharedStringId < 0)
            {
                throw new UnreachableException(
                    $"Unable to find text '{text}' in shared string table for cell {xlCell.Point}. "
                        + "That likely means reference counting is broken. As a stop-gap, try to set the "
                        + "text value to an unused cell to increase number of references for the text."
                );
            }

            return sharedStringId;
        }

        /// <summary>
        /// Get id of number format that is going to be actually saved to database.
        /// </summary>
        internal int? GetNumberFormatId(string? numberFormat)
        {
            if (numberFormat is null)
            {
                return null;
            }

            return this.NumberFormatMap[numberFormat];
        }

        internal int GetFontId(XLFontFormatValue font) => this.FontMap[font];

        internal uint GetDxfId(XLDxfValue dxf) => this.DxfMap[dxf];

        internal uint GetStyleId(XLCellFormatValue? format) =>
            format is not null ? this.FormatMap[format] : 0;
#nullable disable
    }

    #endregion Nested type: SaveContext

    #region Nested type: RelType

    internal enum RelType
    {
        Workbook //, Worksheet
    }

    #endregion Nested type: RelType

    #region Nested type: RelIdGenerator

    internal sealed class RelIdGenerator
    {
        private readonly Dictionary<RelType, HashSet<string>> _relIds = new();

        public void AddValues(IEnumerable<string> values, RelType relType)
        {
            if (!this._relIds.TryGetValue(relType, out HashSet<string> set))
            {
                set = [];
                this._relIds.Add(relType, set);
            }

            set.UnionWith(values);
        }

        /// <summary>
        /// Add all existing rel ids present on the parts or workbook to the generator, so they are not duplicated again.
        /// </summary>
        public void AddExistingValues(WorkbookPart workbookPart, XLWorkbook xlWorkbook)
        {
            this.AddValues(workbookPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
            this.AddValues(
                xlWorkbook
                    .WorksheetsInternal.Cast<XLWorksheet>()
                    .Where(ws => !string.IsNullOrWhiteSpace(ws.RelId))
                    .Select(ws => ws.RelId),
                RelType.Workbook
            );
            this.AddValues(
                xlWorkbook
                    .WorksheetsInternal.Cast<XLWorksheet>()
                    .Where(ws => !string.IsNullOrWhiteSpace(ws.LegacyDrawingId))
                    .Select(ws => ws.LegacyDrawingId),
                RelType.Workbook
            );
            this.AddValues(
                xlWorkbook
                    .WorksheetsInternal.Cast<XLWorksheet>()
                    .SelectMany(ws => ws.Tables.Cast<XLTable>())
                    .Where(t => !string.IsNullOrWhiteSpace(t.RelId))
                    .Select(t => t.RelId),
                RelType.Workbook
            );

            foreach (XLWorksheet xlWorksheet in xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>())
            {
                // if the worksheet is a new one, it doesn't have RelId yet.
                if (
                    string.IsNullOrEmpty(xlWorksheet.RelId)
                    || !workbookPart.TryGetPartById(xlWorksheet.RelId, out OpenXmlPart part)
                )
                {
                    continue;
                }

                WorksheetPart worksheetPart = (WorksheetPart)part;
                this.AddValues(
                    worksheetPart.HyperlinkRelationships.Select(hr => hr.Id),
                    RelType.Workbook
                );
                this.AddValues(worksheetPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
                if (worksheetPart.DrawingsPart != null)
                {
                    this.AddValues(
                        worksheetPart.DrawingsPart.Parts.Select(p => p.RelationshipId),
                        RelType.Workbook
                    );
                }
            }
        }

        public string GetNext(RelType relType)
        {
            if (!this._relIds.TryGetValue(relType, out HashSet<string> set))
            {
                set = [];
                this._relIds.Add(relType, set);
            }

            int id = set.Count + 1;
            while (true)
            {
                string relId = string.Concat("rId", id);
                if (!set.Contains(relId))
                {
                    set.Add(relId);
                    return relId;
                }
                id++;
            }
        }
    }

    #endregion Nested type: RelIdGenerator
}
