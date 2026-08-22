using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Tables;

[DebuggerDisplay("{Name}")]
internal class XLTableField : IXLTableField
{
    internal XLTotalsRowFunction totalsRowFunction;
    internal string? totalsRowLabel;
    private readonly XLTable table;

    private IXLRangeColumn? _column;
    private int index;
    private string name;

    public XLTableField(XLTable table, string name)
    {
        this.table = table;
        this.name = name;
    }

    public IXLRangeColumn Column
    {
        get
        {
            if (this._column == null)
            {
                this._column = this.table.AsRange().Column(this.Index + 1);
            }
            return this._column;
        }
        internal set => this._column = value;
    }

    public IXLCells DataCells =>
        this.Column.Cells(c =>
        {
            if (this.table.ShowHeaderRow && c.Equals(this.HeaderCell))
            {
                return false;
            }

            if (this.table.ShowTotalsRow && c.Equals(this.TotalsCell))
            {
                return false;
            }

            return true;
        });

    public IXLCell? HeaderCell
    {
        get
        {
            if (!this.table.ShowHeaderRow)
            {
                return null;
            }

            return this.Column.FirstCell();
        }
    }

    public int Index
    {
        get => this.index;
        internal set
        {
            if (this.index == value)
            {
                return;
            }

            this.index = value;
            this._column = null;
        }
    }

    public string Name
    {
        get => this.name;
        set
        {
            if (this.name == value)
            {
                return;
            }

            if (this.table.ShowHeaderRow)
            {
                ((XLCell)this.table.HeadersRow(false).Cell(this.Index + 1)).SetValue(
                    value,
                    setTableHeader: false,
                    checkMergedRanges: true
                );
            }

            this.table.RenameField(this.name, value);
            this.name = value;
        }
    }

    public IXLTable Table => this.table;

    public IXLCell? TotalsCell
    {
        get
        {
            if (!this.table.ShowTotalsRow)
            {
                return null;
            }

            return this.Column.LastCell();
        }
    }

    public string TotalsRowFormulaA1
    {
        get => this.table.TotalsRow().Cell(this.Index + 1).FormulaA1;
        set
        {
            this.totalsRowFunction = XLTotalsRowFunction.Custom;
            this.table.TotalsRow().Cell(this.Index + 1).FormulaA1 = value;
        }
    }

    public string TotalsRowFormulaR1C1
    {
        get => this.table.TotalsRow().Cell(this.Index + 1).FormulaR1C1;
        set
        {
            this.totalsRowFunction = XLTotalsRowFunction.Custom;
            this.table.TotalsRow().Cell(this.Index + 1).FormulaR1C1 = value;
        }
    }

    public XLTotalsRowFunction TotalsRowFunction
    {
        get => this.totalsRowFunction;
        set
        {
            this.totalsRowFunction = value;
            this.UpdateTableFieldTotalsRowFormula();
        }
    }

    public string? TotalsRowLabel
    {
        get => this.totalsRowLabel;
        set
        {
            this.totalsRowFunction = XLTotalsRowFunction.None;
            ((XLCell)this.table.TotalsRow().Cell(this.Index + 1)).SetValue(
                value,
                setTableHeader: false,
                checkMergedRanges: true
            );
            this.totalsRowLabel = value;
        }
    }

    /// <summary>
    /// Dxf of header row cells of the fields column.
    /// </summary>
    internal XLDxfValue? HeaderFormatValue { get; set; }

    /// <summary>
    /// Dxf of total data cells of the fields column.
    /// </summary>
    internal XLDxfValue? DataFormatValue { get; set; }

    /// <summary>
    /// Dxf of total row cells of the fields column.
    /// </summary>
    internal XLDxfValue? TotalFormatValue { get; set; }

    public void Delete() => this.Delete(true);

    internal void Delete(bool deleteUnderlyingRangeColumn)
    {
        XLTableField[] fields = [.. this.table.Fields.Cast<XLTableField>()];

        if (deleteUnderlyingRangeColumn)
        {
            this.table.AsRange().ColumnQuick(this.Index + 1).Delete();
        }

        fields.Where(f => f.Index > this.Index).ForEach(f => f.Index--);
        this.table.FieldNames.Remove(this.Name);
    }

    public bool IsConsistentDataType()
    {
        IEnumerable<XLDataType> dataTypes = this
            .Column.Cells()
            .Skip(this.table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.DataType);

        if (this.table.ShowTotalsRow)
        {
            dataTypes = dataTypes.SkipLast();
        }

        var distinctDataTypes = dataTypes
            .GroupBy(dt => dt)
            .Select(g => new { Key = g.Key, Count = g.Count() });

        return distinctDataTypes.Count() == 1;
    }

    public bool IsConsistentFormula()
    {
        IEnumerable<string> formulas = this
            .Column.Cells()
            .Skip(this.table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.FormulaR1C1);

        if (this.table.ShowTotalsRow)
        {
            formulas = formulas.SkipLast();
        }

        var distinctFormulas = formulas
            .GroupBy(f => f)
            .Select(g => new { Key = g.Key, Count = g.Count() });

        return distinctFormulas.Count() == 1;
    }

    public bool IsConsistentStyle()
    {
        IEnumerable<XLCellFormat> styles = this
            .Column.Cells()
            .Skip(this.table.ShowHeaderRow ? 1 : 0)
            .OfType<XLCell>()
            .Select(c => c.Format);

        if (this.table.ShowTotalsRow)
        {
            styles = styles.SkipLast();
        }

        IEnumerable<XLCellFormat> distinctStyles = styles.Distinct();

        return distinctStyles.Count() == 1;
    }

    private static IEnumerable<string> QuotedTableFieldCharacters = new[] { "'", "#" };

    internal void UpdateTableFieldTotalsRowFormula()
    {
        if (
            this.TotalsRowFunction != XLTotalsRowFunction.None
            && this.TotalsRowFunction != XLTotalsRowFunction.Custom
        )
        {
            IXLCell cell = this.table.TotalsRow().Cell(this.Index + 1);
            string formulaCode = string.Empty;
            switch (this.TotalsRowFunction)
            {
                case XLTotalsRowFunction.Sum:
                    formulaCode = "109";
                    break;
                case XLTotalsRowFunction.Minimum:
                    formulaCode = "105";
                    break;
                case XLTotalsRowFunction.Maximum:
                    formulaCode = "104";
                    break;
                case XLTotalsRowFunction.Average:
                    formulaCode = "101";
                    break;
                case XLTotalsRowFunction.Count:
                    formulaCode = "103";
                    break;
                case XLTotalsRowFunction.CountNumbers:
                    formulaCode = "102";
                    break;
                case XLTotalsRowFunction.StandardDeviation:
                    formulaCode = "107";
                    break;
                case XLTotalsRowFunction.Variance:
                    formulaCode = "110";
                    break;
            }

            string modifiedName = this.Name;
            QuotedTableFieldCharacters.ForEach(c =>
                modifiedName = modifiedName.Replace(c, "'" + c)
            );

            if (modifiedName.StartsWith(" ") || modifiedName.EndsWith(" "))
            {
                modifiedName = "[" + modifiedName + "]";
            }

            bool prependTableName = modifiedName.Contains(" ");

            cell.FormulaA1 =
                $"SUBTOTAL({formulaCode},{(prependTableName ? this.table.Name : string.Empty)}[{modifiedName}])";
            IXLCell lastCell = this.table.LastRow().Cell(this.Index + 1);
            if (lastCell.DataType != XLDataType.Text)
            {
                cell.Style.NumberFormat = lastCell.Style.NumberFormat;
            }
        }
    }
}
