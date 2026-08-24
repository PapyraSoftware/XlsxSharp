#nullable disable

using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Text;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Tables;

[DebuggerDisplay("{Name}")]
internal class XLTable : XLRange, IXLTable
{
    private string _name;
    internal bool _showTotalsRow;
    internal HashSet<string> _uniqueNames;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLTable(XLRangeAddress tableArea, IXLStyle defaultStyle)
        : base(tableArea, defaultStyle) => this.InitializeValues(false);

    public override XLRangeType RangeType => XLRangeType.Table;

    private IXLRangeAddress _lastRangeAddress;
    private Dictionary<string, XLTableField> _fieldNames = CreateFieldNames();

    internal Dictionary<string, XLTableField> FieldNames
    {
        get
        {
            if (
                this._fieldNames != null
                && this._lastRangeAddress != null
                && this._lastRangeAddress.Equals(this.RangeAddress)
            )
            {
                return this._fieldNames;
            }

            this._lastRangeAddress = this.RangeAddress;

            this.RescanFieldNames();

            return this._fieldNames;
        }
    }

    /// <summary>
    /// Area of the range, including headings and totals, if table has them.
    /// </summary>
    internal Area Area => Area.FromRangeAddress(this.RangeAddress);

    private void RescanFieldNames()
    {
        if (this.ShowHeaderRow)
        {
            Dictionary<string, XLTableField> oldFieldNames = this._fieldNames ?? CreateFieldNames();
            this._fieldNames = CreateFieldNames();
            XLRangeRow headersRow = this.HeadersRow(false);
            int cellPos = 0;
            foreach (XLCell cell in headersRow.Cells())
            {
                XLCellValue cellValue = cell.CachedValue;
                string name = cellValue.ToString(CultureInfo.CurrentCulture);

                if (oldFieldNames.TryGetValue(name, out XLTableField tableField)) // && tableField.Column.ColumnNumber() == cell.Address.ColumnNumber)
                {
                    tableField.Index = cellPos;
                    this._fieldNames.Add(name, tableField);
                    cellPos++;
                    continue;
                }

                // Be careful here. Fields names may actually be whitespace, but not empty
                if (string.IsNullOrEmpty(name))
                {
                    name = this.GetUniqueName("Column", cellPos + 1, true);
                }
                if (this._fieldNames.ContainsKey(name))
                {
                    throw new ArgumentException(
                        "The header row contains more than one field name '" + name + "'."
                    );
                }

                this._fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });

                // Field names are the source of the truth that is projected
                // to the cells and field names can be only text. Fix the cell,
                // so cell fulfills its job of being dependent on the field name.
                if (!cellValue.Equals(name))
                {
                    cell.SetValue(name, false, false);
                }
            }
        }
        else
        {
            int colCount = this.ColumnCount();
            for (int i = 1; i <= colCount; i++)
            {
                if (this._fieldNames.Values.All(f => f.Index != i - 1))
                {
                    string name = "Column" + i;

                    this._fieldNames.Add(name, new XLTableField(this, name) { Index = i - 1 });
                }
            }
        }
    }

    internal XLTableField AddField(string fieldName)
    {
        XLTableField field = new(this, fieldName) { Index = this._fieldNames.Count };
        this._fieldNames.Add(fieldName, field);
        return field;
    }

    internal void RenameField(string oldName, string newName)
    {
        if (!this._fieldNames.TryGetValue(oldName, out XLTableField field))
        {
            throw new ArgumentException("The field does not exist in this table", "oldName");
        }

        this._fieldNames.Remove(oldName);
        this._fieldNames.Add(newName, field);
    }

    internal string RelId { get; set; }

    public IXLTableRange DataRange
    {
        get
        {
            XLRange range;

            int firstDataRowNumber = 1;
            int lastDataRowNumber = this.RowCount();

            if (this._showHeaderRow)
            {
                firstDataRowNumber++;
            }

            if (this._showTotalsRow)
            {
                lastDataRowNumber--;
            }

            if (firstDataRowNumber > lastDataRowNumber)
            {
                return null;
            }

            range = this.Range(firstDataRowNumber, 1, lastDataRowNumber, this.ColumnCount());

            return new XLTableRange(range, this);
        }
    }

    private XLAutoFilter _autoFilter;

    public XLAutoFilter AutoFilter
    {
        get
        {
            if (this._autoFilter == null)
            {
                this._autoFilter = new XLAutoFilter();
            }

            this._autoFilter.Range = this.ShowTotalsRow
                ? this.Range(1, 1, this.RowCount() - 1, this.ColumnCount())
                : this.AsRange();
            return this._autoFilter;
        }
    }

    public override IXLAutoFilter SetAutoFilter() => this.AutoFilter;

    protected override void OnRangeAddressChanged(
        XLRangeAddress oldAddress,
        XLRangeAddress newAddress
    )
    {
        //Do nothing for table
    }

    #region IXLTable Members

    public bool EmphasizeFirstColumn { get; set; }
    public bool EmphasizeLastColumn { get; set; }
    public bool ShowRowStripes { get; set; }
    public bool ShowColumnStripes { get; set; }

    private bool _showAutoFilter;

    public bool ShowAutoFilter
    {
        get => this._showHeaderRow && this._showAutoFilter;
        set => this._showAutoFilter = value;
    }

    public XLTableTheme Theme { get; set; }

    public string Name
    {
        get => this._name;
        set
        {
            if (this._name == value)
            {
                return;
            }

            // Validation rules for table names
            string oldname = this._name ?? string.Empty;
            IEnumerable<string> tableNames = this.Worksheet.Tables.Select<XLTable, string>(t =>
                t.Name
            );
            if (
                !XlsxSharp.XLHelper.ValidateName(
                    "table",
                    value,
                    oldname,
                    tableNames,
                    out string message
                )
            )
            {
                throw new ArgumentException(message, nameof(value));
            }

            this._name = value;

            // Some totals row formula depend on the table name. Update them.
            if (this._fieldNames?.Any() ?? false)
            {
                this.Fields.ForEach(f => (f as XLTableField).UpdateTableFieldTotalsRowFormula());
            }

            if (
                !string.IsNullOrWhiteSpace(oldname)
                && !string.Equals(oldname, this._name, StringComparison.OrdinalIgnoreCase)
            )
            {
                this.Worksheet.Tables.Add(this);
                if (this.Worksheet.Tables.Contains(oldname))
                {
                    this.Worksheet.Tables.Remove(oldname);
                }
            }
        }
    }

    public bool ShowTotalsRow
    {
        get => this._showTotalsRow;
        set
        {
            if (value && !this._showTotalsRow)
            {
                this.InsertRowsBelow(1);
            }
            else if (!value && this._showTotalsRow)
            {
                this.TotalsRow().Delete();
            }

            this._showTotalsRow = value;

            // Invalidate fields' columns
            this.Fields.Cast<XLTableField>().ForEach(f => f.Column = null);

            if (this._showTotalsRow)
            {
                this.AutoFilter.Range = this.Worksheet.Range(
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.RowNumber - 1,
                    this.RangeAddress.LastAddress.ColumnNumber
                );
            }
            else
            {
                this.AutoFilter.Range = this.Worksheet.Range(this.RangeAddress);
            }
        }
    }

    public IXLRangeRow HeadersRow() => this.HeadersRow(true);

    internal XLRangeRow HeadersRow(bool scanForNewFieldsNames)
    {
        if (!this.ShowHeaderRow)
        {
            return null;
        }

        if (scanForNewFieldsNames)
        {
            Dictionary<string, XLTableField> tempResult = this.FieldNames;
        }

        return this.FirstRow();
    }

    public IXLRangeRow TotalsRow() => this.ShowTotalsRow ? this.LastRow() : null;

    public IXLTableField Field(string fieldName) => this.Field(this.GetFieldIndex(fieldName));

    IXLTableField IXLTable.Field(int fieldIndex) => this.Field(fieldIndex);

    internal XLTableField Field(int fieldIndex) =>
        this.FieldNames.Values.First(f => f.Index == fieldIndex);

    IEnumerable<IXLTableField> IXLTable.Fields => this.Fields;

    internal IEnumerable<XLTableField> Fields
    {
        get
        {
            int columnCount = this.ColumnCount();
            for (int co = 0; co < columnCount; co++)
            {
                yield return this.Field(co);
            }
        }
    }

    public IXLTable Resize(IXLRangeAddress rangeAddress) =>
        this.Resize(this.Worksheet.Range(this.RangeAddress));

    public IXLTable Resize(string rangeAddress) => this.Resize(this.Worksheet.Range(rangeAddress));

    public IXLTable Resize(IXLCell firstCell, IXLCell lastCell) =>
        this.Resize(this.Worksheet.Range(firstCell, lastCell));

    public IXLTable Resize(string firstCellAddress, string lastCellAddress) =>
        this.Resize(this.Worksheet.Range(firstCellAddress, lastCellAddress));

    public IXLTable Resize(IXLAddress firstCellAddress, IXLAddress lastCellAddress) =>
        this.Resize(this.Worksheet.Range(firstCellAddress, lastCellAddress));

    public IXLTable Resize(
        int firstCellRow,
        int firstCellColumn,
        int lastCellRow,
        int lastCellColumn
    ) =>
        this.Resize(
            this.Worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn)
        );

    public IXLTable Resize(IXLRange range)
    {
        if (!this.ShowHeaderRow)
        {
            throw new NotImplementedException(
                "Resizing of tables with no headers not supported yet."
            );
        }

        if (this.Worksheet != range.Worksheet)
        {
            throw new InvalidOperationException(
                "You cannot resize a table to a range on a different sheet."
            );
        }

        int totalsRowChanged = this.ShowTotalsRow
            ? range.LastRow().RowNumber() - this.TotalsRow().RowNumber()
            : 0;
        int oldTotalsRowNumber = this.ShowTotalsRow ? this.TotalsRow().RowNumber() : -1;

        Dictionary<string, XLTableField>.KeyCollection existingHeaders = this.FieldNames.Keys;
        HashSet<string> newHeaders = [];

        // Force evaluation of f.Column field
        IXLRangeColumn[] tempArray = [.. this.Fields.Select(f => f.Column)];

        IXLRangeRow firstRow = range.Row(1);
        if (
            !firstRow.FirstCell().Address.Equals(this.HeadersRow().FirstCell().Address)
            || !firstRow.LastCell().Address.Equals(this.HeadersRow().LastCell().Address)
        )
        {
            this._uniqueNames.Clear();
            int co = 1;
            foreach (IXLCell c in firstRow.Cells())
            {
                if (c.IsEmpty(XLCellsUsedOptions.Contents))
                {
                    c.Value = this.GetUniqueName("Column", co, true);
                }

                string header = c.GetString();
                this._uniqueNames.Add(header);

                if (!existingHeaders.Contains(header))
                {
                    newHeaders.Add(header);
                }

                co++;
            }
        }

        if (totalsRowChanged < 0)
        {
            range
                .Rows(r => r.RowNumber().Equals(this.TotalsRow().RowNumber() + totalsRowChanged))
                .Single()
                .InsertRowsAbove(1);
            range = this.Worksheet.Range(range.FirstCell(), range.LastCell().CellAbove());
            oldTotalsRowNumber++;
        }
        else if (totalsRowChanged > 0)
        {
            this.TotalsRow().RowBelow(totalsRowChanged + 1).InsertRowsAbove(1);
            this.TotalsRow().AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        this.RangeAddress = (XLRangeAddress)range.RangeAddress;
        this.RescanFieldNames();

        if (this.ShowTotalsRow)
        {
            foreach (XLTableField f in this._fieldNames.Values)
            {
                int fieldColumn = f.Index + 1;
                IXLCell c = this.TotalsRow().Cell(fieldColumn);
                if (!c.IsEmpty() && newHeaders.Contains(f.Name))
                {
                    f.TotalsRowLabel = c.GetFormattedString();
                }
            }

            if (totalsRowChanged != 0)
            {
                foreach (XLTableField f in this._fieldNames.Values.Cast<XLTableField>())
                {
                    f.UpdateTableFieldTotalsRowFormula();
                    int fieldColumn = f.Index + 1;
                    IXLCell c = this.TotalsRow().Cell(fieldColumn);
                    if (!string.IsNullOrWhiteSpace(f.TotalsRowLabel))
                    {
                        //Remove previous row's label
                        XLCell oldTotalsCell = this.Worksheet.Cell(
                            oldTotalsRowNumber,
                            f.Column.ColumnNumber()
                        );
                        if (oldTotalsCell.Value.Equals(f.TotalsRowLabel))
                        {
                            oldTotalsCell.Value = Blank.Value;
                        }
                    }

                    if (!string.IsNullOrEmpty(f.TotalsRowLabel))
                    {
                        c.SetValue(f.TotalsRowLabel);
                    }
                }
            }
        }

        return this;
    }

    public IXLTable SetEmphasizeFirstColumn()
    {
        this.EmphasizeFirstColumn = true;
        return this;
    }

    public IXLTable SetEmphasizeFirstColumn(bool value)
    {
        this.EmphasizeFirstColumn = value;
        return this;
    }

    public IXLTable SetEmphasizeLastColumn()
    {
        this.EmphasizeLastColumn = true;
        return this;
    }

    public IXLTable SetEmphasizeLastColumn(bool value)
    {
        this.EmphasizeLastColumn = value;
        return this;
    }

    public IXLTable SetShowRowStripes()
    {
        this.ShowRowStripes = true;
        return this;
    }

    public IXLTable SetShowRowStripes(bool value)
    {
        this.ShowRowStripes = value;
        return this;
    }

    public IXLTable SetShowColumnStripes()
    {
        this.ShowColumnStripes = true;
        return this;
    }

    public IXLTable SetShowColumnStripes(bool value)
    {
        this.ShowColumnStripes = value;
        return this;
    }

    public IXLTable SetShowTotalsRow()
    {
        this.ShowTotalsRow = true;
        return this;
    }

    public IXLTable SetShowTotalsRow(bool value)
    {
        this.ShowTotalsRow = value;
        return this;
    }

    public IXLTable SetShowAutoFilter()
    {
        this.ShowAutoFilter = true;
        return this;
    }

    public IXLTable SetShowAutoFilter(bool value)
    {
        this.ShowAutoFilter = value;
        return this;
    }

    public new IXLRange Sort(
        string columnsToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    )
    {
        StringBuilder toSortBy = new();
        foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
        {
            string coString;
            string order;
            if (coPairTrimmed.Contains(' '))
            {
                string[] pair = coPairTrimmed.Split(' ');
                coString = pair[0];
                order = pair[1];
            }
            else
            {
                coString = coPairTrimmed;
                order = sortOrder == XLSortOrder.Ascending ? "ASC" : "DESC";
            }

            if (!int.TryParse(coString, out int co))
            {
                co = this.Field(coString).Index + 1;
            }

            if (toSortBy.Length > 0)
            {
                toSortBy.Append(',');
            }

            toSortBy.Append(co);
            toSortBy.Append(' ');
            toSortBy.Append(order);
        }
        return this.DataRange.Sort(toSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
    }

    public new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    IXLAutoFilter IXLTable.AutoFilter => this.AutoFilter;

    #endregion IXLTable Members

    private void InitializeValues(bool setAutofilter)
    {
        this.ShowRowStripes = true;
        this._showHeaderRow = true;
        this.Theme = XLTableTheme.TableStyleMedium2;
        if (setAutofilter)
        {
            this.InitializeAutoFilter();
        }

        if (this.RowCount() == 1)
        {
            this.InsertRowsBelow(1);
        }
    }

    public void InitializeAutoFilter() => this.ShowAutoFilter = true;

    internal void OnAddedToTables()
    {
        this._uniqueNames = [];
        int co = 1;
        foreach (IXLCell c in this.Row(1).Cells())
        {
            // Be careful here. Fields names may actually be whitespace, but not empty
            if (c.IsEmpty(XLCellsUsedOptions.Contents))
            {
                (c as XLCell).SetValue(this.GetUniqueName("Column", co, true), false, false);
            }

            this._uniqueNames.Add(c.GetString());
            co++;
        }
    }

    private static Dictionary<string, XLTableField> CreateFieldNames() =>
        new(StringComparer.CurrentCultureIgnoreCase);

    private string GetUniqueName(string originalName, int initialOffset, bool enforceOffset)
    {
        string name = string.Concat(
            originalName,
            enforceOffset ? initialOffset.ToInvariantString() : string.Empty
        );
        if (this._uniqueNames?.Contains(name) ?? false)
        {
            int i = initialOffset;
            name = originalName + i.ToInvariantString();
            while (this._uniqueNames.Contains(name))
            {
                i++;
                name = originalName + i.ToInvariantString();
            }
        }

        return name;
    }

    public int GetFieldIndex(string name)
    {
        // There is a discrepancy in the way headers with line breaks are stored.
        // The entry in the table definition will contain \r\n
        // but the shared string value of the actual cell will contain only \n
        name = name.Replace("\r\n", "\n");
        if (this.FieldNames.TryGetValue(name, out XLTableField tableField))
        {
            return tableField.Index;
        }

        throw new ArgumentOutOfRangeException(
            "The header row doesn't contain field name '" + name + "'."
        );
    }

    internal bool _showHeaderRow;

    public bool ShowHeaderRow
    {
        get => this._showHeaderRow;
        set
        {
            if (this._showHeaderRow == value)
            {
                return;
            }

            if (this._showHeaderRow)
            {
                IXLRangeRow headersRow = this.HeadersRow();
                this._uniqueNames = [];
                int co = 1;
                foreach (IXLCell c in headersRow.Cells())
                {
                    if (string.IsNullOrWhiteSpace(c.GetString()))
                    {
                        c.Value = this.GetUniqueName("Column", co, true);
                    }

                    this._uniqueNames.Add(c.GetString());
                    co++;
                }

                headersRow.Clear();
                this.RangeAddress = new XLRangeAddress(
                    new XLAddress(
                        this.Worksheet,
                        this.RangeAddress.FirstAddress.RowNumber + 1,
                        this.RangeAddress.FirstAddress.ColumnNumber,
                        this.RangeAddress.FirstAddress.FixedRow,
                        this.RangeAddress.FirstAddress.FixedColumn
                    ),
                    this.RangeAddress.LastAddress
                );
            }
            else
            {
                XLRange asRange = this.Worksheet.Range(
                    this.RangeAddress.FirstAddress.RowNumber - 1,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber
                );
                XLRangeRow firstRow = asRange.FirstRow();
                IXLRangeRow rangeRow;
                if (firstRow.IsEmpty(XLCellsUsedOptions.All))
                {
                    rangeRow = firstRow;
                    this.RangeAddress = new XLRangeAddress(
                        new XLAddress(
                            this.Worksheet,
                            this.RangeAddress.FirstAddress.RowNumber - 1,
                            this.RangeAddress.FirstAddress.ColumnNumber,
                            this.RangeAddress.FirstAddress.FixedRow,
                            this.RangeAddress.FirstAddress.FixedColumn
                        ),
                        this.RangeAddress.LastAddress
                    );
                }
                else
                {
                    XLAddress fAddress = this.RangeAddress.FirstAddress;
                    //var lAddress = RangeAddress.LastAddress;

                    rangeRow = firstRow.InsertRowsBelow(1, false).First();

                    this.RangeAddress = new XLRangeAddress(fAddress, this.RangeAddress.LastAddress);
                }

                int co = 1;
                foreach (string name in this.FieldNames.Values.Select(f => f.Name))
                {
                    rangeRow.Cell(co).SetValue(name);
                    co++;
                }
            }

            this._showHeaderRow = value;

            // Invalidate fields' columns
            this.Fields.Cast<XLTableField>().ForEach(f => f.Column = null);
        }
    }

    public IXLTable SetShowHeaderRow() => this.SetShowHeaderRow(true);

    public IXLTable SetShowHeaderRow(bool value)
    {
        this.ShowHeaderRow = value;
        return this;
    }

    public void ExpandTableRows(int rows) =>
        this.RangeAddress = new XLRangeAddress(
            this.RangeAddress.FirstAddress,
            new XLAddress(
                this.Worksheet,
                this.RangeAddress.LastAddress.RowNumber + rows,
                this.RangeAddress.LastAddress.ColumnNumber,
                this.RangeAddress.LastAddress.FixedRow,
                this.RangeAddress.LastAddress.FixedColumn
            )
        );

    public override XLRangeColumn Column(int columnNumber)
    {
        XLRangeColumn column = base.Column(columnNumber);
        column.Table = this;
        return column;
    }

    public override XLRangeColumn Column(string columnName)
    {
        XLRangeColumn column = base.Column(columnName);
        column.Table = this;
        return column;
    }

    public override IXLRangeColumns Columns(int firstColumn, int lastColumn)
    {
        IXLRangeColumns columns = base.Columns(firstColumn, lastColumn);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    public override IXLRangeColumns Columns(Func<IXLRangeColumn, bool> predicate = null)
    {
        IXLRangeColumns columns = base.Columns(predicate);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    public override IXLRangeColumns Columns(string columns)
    {
        IXLRangeColumns cols = base.Columns(columns);
        cols.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return cols;
    }

    public override IXLRangeColumns Columns(string firstColumn, string lastColumn)
    {
        IXLRangeColumns columns = base.Columns(firstColumn, lastColumn);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    internal override XLRangeColumns ColumnsUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, bool> predicate = null
    )
    {
        XLRangeColumns columns = base.ColumnsUsed(options, predicate);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    internal override XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, bool> predicate = null)
    {
        XLRangeColumns columns = base.ColumnsUsed(predicate);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, string name) =>
        this.CreatePivotTable(targetCell, name);

    internal new XLPivotTable CreatePivotTable(IXLCell targetCell, string name) =>
        (XLPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, this);

    public IEnumerable<dynamic> AsDynamicEnumerable()
    {
        foreach (IXLTableRow row in this.DataRange.Rows())
        {
            dynamic expando = new ExpandoObject();
            foreach (XLTableField f in this.Fields)
            {
                XLCellValue value = row.Cell(f.Index + 1).Value;
                // ExpandoObject supports IDictionary so we can extend it like this
                IDictionary<string, object> expandoDict = expando as IDictionary<string, object>;
                expandoDict[f.Name] = value;
            }

            yield return expando;
        }
    }

    public DataTable AsNativeDataTable()
    {
        DataTable table = new(this.Name);

        foreach (XLTableField f in this.Fields.Cast<XLTableField>())
        {
            Type type = typeof(object);
            if (f.IsConsistentDataType())
            {
                IXLCell c = f.Column.Cells().Skip(this.ShowHeaderRow ? 1 : 0).First();
                switch (c.DataType)
                {
                    case XLDataType.Text:
                        type = typeof(string);
                        break;

                    case XLDataType.Boolean:
                        type = typeof(bool);
                        break;

                    case XLDataType.DateTime:
                        type = typeof(DateTime);
                        break;

                    case XLDataType.TimeSpan:
                        type = typeof(TimeSpan);
                        break;

                    case XLDataType.Number:
                        type = typeof(double);
                        break;
                }
            }

            table.Columns.Add(f.Name, type);
        }

        foreach (IXLTableRow row in this.DataRange.Rows())
        {
            DataRow dr = table.NewRow();

            foreach (XLTableField f in this.Fields)
            {
                dr[f.Name] = row.Cell(f.Index + 1).Value.ToObject();
            }

            table.Rows.Add(dr);
        }

        return table;
    }

    public IXLTable CopyTo(IXLWorksheet targetSheet) => this.CopyTo((XLWorksheet)targetSheet);

    internal IXLTable CopyTo(XLWorksheet targetSheet, bool copyData = true)
    {
        if (targetSheet == this.Worksheet)
        {
            throw new InvalidOperationException(
                "Cannot copy table to the worksheet it already belongs to."
            );
        }

        XLRange targetRange = targetSheet.Range(this.RangeAddress.WithoutWorksheet());
        if (copyData)
        {
            this.RangeUsed().CopyTo(targetRange);
        }
        else
        {
            this.HeadersRow().CopyTo(targetRange.FirstRow());
        }

        string tableName = this.Name;
        XLTable newTable = (XLTable)targetSheet.Table(targetRange, tableName, true);

        newTable.RelId = null;
        newTable.EmphasizeFirstColumn = this.EmphasizeFirstColumn;
        newTable.EmphasizeLastColumn = this.EmphasizeLastColumn;
        newTable.ShowRowStripes = this.ShowRowStripes;
        newTable.ShowColumnStripes = this.ShowColumnStripes;
        newTable.ShowAutoFilter = this.ShowAutoFilter;
        newTable.Theme = this.Theme;
        newTable._showTotalsRow = this.ShowTotalsRow;

        int fieldCount = this.ColumnCount();
        for (int f = 0; f < fieldCount; f++)
        {
            XLTableField tableField = newTable.Field(f) as XLTableField;
            XLTableField tField = this.Field(f) as XLTableField;
            tableField.Index = tField.Index;
            tableField.Name = tField.Name;
            tableField.totalsRowLabel = tField.totalsRowLabel;
            tableField.totalsRowFunction = tField.totalsRowFunction;
        }
        return newTable;
    }

    #region Append and replace data

    public IXLRange AppendData(IEnumerable data, bool propagateExtraColumns = false) =>
        this.AppendData(data, transpose: false, propagateExtraColumns: propagateExtraColumns);

    public IXLRange AppendData(IEnumerable data, bool transpose, bool propagateExtraColumns = false)
    {
        object[] castedData = data?.Cast<object>().ToArray() ?? [];
        if (!castedData.Any() || data is string)
        {
            return null;
        }

        int numberOfNewRows = castedData.Length;

        IXLTableRow lastRowOfOldRange = this.DataRange.LastRow();
        lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);
        this.Fields.Cast<XLTableField>().ForEach(f => f.Column = null);

        IXLRange insertedRange = lastRowOfOldRange
            .RowBelow()
            .FirstCell()
            .InsertData(castedData, transpose);

        this.PropagateExtraColumns(insertedRange.ColumnCount(), lastRowOfOldRange.RowNumber());

        return insertedRange;
    }

    public IXLRange AppendData(DataTable dataTable, bool propagateExtraColumns = false) =>
        this.AppendData(
            dataTable.Rows.Cast<DataRow>(),
            propagateExtraColumns: propagateExtraColumns
        );

    public IXLRange AppendData<T>(IEnumerable<T> data, bool propagateExtraColumns = false)
    {
        T[] materializedData = data?.ToArray() ?? [];
        if (!materializedData.Any() || data is string)
        {
            return null;
        }

        int numberOfNewRows = materializedData.Length;

        if (numberOfNewRows == 0)
        {
            return null;
        }

        IXLTableRow lastRowOfOldRange = this.DataRange.LastRow();
        lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);
        this.Fields.Cast<XLTableField>().ForEach(f => f.Column = null);

        IXLRange insertedRange = lastRowOfOldRange
            .RowBelow()
            .FirstCell()
            .InsertData(materializedData);

        this.PropagateExtraColumns(insertedRange.ColumnCount(), lastRowOfOldRange.RowNumber());

        return insertedRange;
    }

    public IXLRange ReplaceData(IEnumerable data, bool propagateExtraColumns = false) =>
        this.ReplaceData(data, transpose: false, propagateExtraColumns: propagateExtraColumns);

    public IXLRange ReplaceData(
        IEnumerable data,
        bool transpose,
        bool propagateExtraColumns = false
    )
    {
        object[] castedData = data?.Cast<object>().ToArray() ?? [];
        if (!castedData.Any() || data is string)
        {
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");
        }

        int firstDataRowNumber = this.DataRange.FirstRow().RowNumber();
        int lastDataRowNumber = this.DataRange.LastRow().RowNumber();

        // Resize table
        int sizeDifference = castedData.Length - this.DataRange.RowCount();
        if (sizeDifference > 0)
        {
            this.DataRange.LastRow().InsertRowsBelow(sizeDifference);
        }
        else if (sizeDifference < 0)
        {
            this.DataRange.Rows(
                    lastDataRowNumber + sizeDifference + 1 - firstDataRowNumber + 1,
                    lastDataRowNumber - firstDataRowNumber + 1
                )
                .Delete();

            // No propagation needed when reducing the number of rows
            propagateExtraColumns = false;
        }

        if (sizeDifference != 0)
        // Invalidate table fields' columns
        {
            this.Fields.Cast<XLTableField>().ForEach(f => f.Column = null);
        }

        IXLRange replacedRange = this.DataRange.FirstCell().InsertData(castedData, transpose);

        if (propagateExtraColumns)
        {
            this.PropagateExtraColumns(replacedRange.ColumnCount(), lastDataRowNumber);
        }

        return replacedRange;
    }

    public IXLRange ReplaceData(DataTable dataTable, bool propagateExtraColumns = false) =>
        this.ReplaceData(
            dataTable.Rows.Cast<DataRow>(),
            propagateExtraColumns: propagateExtraColumns
        );

    public IXLRange ReplaceData<T>(IEnumerable<T> data, bool propagateExtraColumns = false)
    {
        T[] materializedData = data?.ToArray() ?? [];
        if (!materializedData.Any() || data is string)
        {
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");
        }

        int firstDataRowNumber = this.DataRange.FirstRow().RowNumber();
        int lastDataRowNumber = this.DataRange.LastRow().RowNumber();

        // Resize table
        int sizeDifference = materializedData.Length - this.DataRange.RowCount();
        if (sizeDifference > 0)
        {
            this.DataRange.LastRow().InsertRowsBelow(sizeDifference);
        }
        else if (sizeDifference < 0)
        {
            this.DataRange.Rows(
                    lastDataRowNumber + sizeDifference + 1 - firstDataRowNumber + 1,
                    lastDataRowNumber - firstDataRowNumber + 1
                )
                .Delete();

            // No propagation needed when reducing the number of rows
            propagateExtraColumns = false;
        }

        if (sizeDifference != 0)
        // Invalidate table fields' columns
        {
            this.Fields.Cast<XLTableField>().ForEach(f => f.Column = null);
        }

        IXLRange replacedRange = this.DataRange.FirstCell().InsertData(materializedData);

        if (propagateExtraColumns)
        {
            this.PropagateExtraColumns(replacedRange.ColumnCount(), lastDataRowNumber);
        }

        return replacedRange;
    }

    private void PropagateExtraColumns(int numberOfNonExtraColumns, int previousLastDataRow)
    {
        for (int i = numberOfNonExtraColumns; i < this.Fields.Count(); i++)
        {
            XLTableField field = this.Field(i);

            XLCell cell = this.Worksheet.Cell(previousLastDataRow, field.Column.ColumnNumber());
            field
                .Column.Cells(c => c.Address.RowNumber > previousLastDataRow)
                .ForEach(c =>
                {
                    if (cell.HasFormula)
                    {
                        c.FormulaR1C1 = cell.FormulaR1C1;
                    }
                    else
                    {
                        c.Value = cell.Value;
                    }
                });
        }
    }

    /// <summary>
    /// Update headers fields and totals fields by data from the cells. Do not add a new fields or names.
    /// </summary>
    /// <param name="refreshArea">Area that contains cells with changed values that might affect header and totals fields.</param>
    internal void RefreshFieldsFromCells(Area refreshArea)
    {
        Area tableArea = this.Area;
        if (this.ShowTotalsRow)
        {
            Area totalsRow = tableArea.SliceFromBottom(1);
            Area? intersection = totalsRow.Intersect(refreshArea);
            if (intersection is not null)
            {
                int totalsRowNumber = totalsRow.BottomRow;
                ValueSlice valueSlice = this.Worksheet.Internals.CellsCollection.ValueSlice;
                for (
                    int column = intersection.Value.LeftColumn;
                    column <= intersection.Value.RightColumn;
                    ++column
                )
                {
                    int fieldIndex = column - totalsRow.LeftColumn;
                    XLTableField field = this.Field(fieldIndex);
                    XLCellValue value = valueSlice.GetCellValue(new Point(totalsRowNumber, column));

                    // Convert value to text, because Excel always converts values to text when replacing totals row.
                    field.TotalsRowLabel = value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }

        if (this.ShowHeaderRow)
        {
            Area headersRow = this.Area.SliceFromTop(1);
            Area? intersection = headersRow.Intersect(refreshArea);
            if (intersection is not null)
            {
                int headersRowNumber = headersRow.TopRow;
                ValueSlice valueSlice = this.Worksheet.Internals.CellsCollection.ValueSlice;
                for (
                    int column = intersection.Value.LeftColumn;
                    column <= intersection.Value.RightColumn;
                    ++column
                )
                {
                    int fieldIndex = column - headersRow.LeftColumn;
                    XLTableField field = this.Field(fieldIndex);
                    XLCellValue value = valueSlice.GetCellValue(
                        new Point(headersRowNumber, column)
                    );

                    // Convert to text, because headers row of a table can be only
                    // string in OOXML and Excel converts it to string as well.
                    field.Name = value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }
    }

    #endregion Append and replace data
}
