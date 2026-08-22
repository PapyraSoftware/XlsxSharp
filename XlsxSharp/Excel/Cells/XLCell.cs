#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Parser;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Excel.Comments;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.InsertData;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.Graphics;

namespace XlsxSharp.Excel;

[DebuggerDisplay("{Address}")]
internal sealed partial class XLCell : IXLCell, IXLFormatContainer
{
    //  @"(?<=\W)" // Start with non word
    [GeneratedRegex(
        @"(?<Reference>" // Start Group to pick
            + @"(?<Sheet>" // Start Sheet Name, optional
            + @"("
            + @"\'([^\[\]\*/\\\?:\']+|\'\')\'"
            // Sheet name with special characters, surrounding apostrophes are required
            + @"|"
            + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
            + @")"
            + @"!)?" // End Sheet Name, optional
            + @"(?<Range>" // Start range
            + @"(?<![\w\d])" // Preceded by anything but a letter or a number
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
            + @"(?![\w\d])" // followed by anything but a letter or a number
            + @"|"
            + @"(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})" // 1:1
            + @"|"
            + @"(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
            + @")" // End Range
            + @")" // End Group to pick
    //+ @"(?=\W)" // End with non word
    )]
    public static partial Regex A1SimpleRegex { get; }

    // 1:1
    [GeneratedRegex(@"(\$?\d{1,7}:\$?\d{1,7})")]
    private static partial Regex A1RowRegex { get; }

    // A:A
    [GeneratedRegex(@"(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})")]
    private static partial Regex A1ColumnRegex { get; }

    [GeneratedRegex(@"(?<!_x005F)_x(?!005F)([0-9A-F]{4})_")]
    private static partial Regex utfPattern { get; }

    private readonly XLCellsCollection _cellsCollection;

    private readonly int _rowNumber;

    private readonly int _columnNumber;

    internal XLCell(XLWorksheet worksheet, int row, int column)
    {
        this._cellsCollection = worksheet.Internals.CellsCollection;
        this._rowNumber = row;
        this._columnNumber = column;
    }

    internal XLCell(XLWorksheet worksheet, Point point)
        : this(worksheet, point.Row, point.Column) { }

    public XLWorksheet Worksheet => this._cellsCollection.Worksheet;

    public XLAddress Address =>
        new(this.Worksheet, this._rowNumber, this._columnNumber, false, false);

    internal Point Point => new(this._rowNumber, this._columnNumber);

    private XLWorkbookStyles Styles => this.Worksheet.Workbook.Styles;

    #region Slice fields

    /// <summary>
    /// A flag indicating if a string should be stored in the shared table or inline.
    /// </summary>
    public bool ShareString
    {
        get => this._cellsCollection.ValueSlice.GetShareString(this.Point);
        set => this._cellsCollection.ValueSlice.SetShareString(this.Point, value);
    }

    internal int MemorySstId => this._cellsCollection.ValueSlice.GetShareStringId(this.Point);

    internal XLImmutableRichText RichText => this.SliceRichText;

    private XLCellValue SliceCellValue
    {
        get => this._cellsCollection.ValueSlice.GetCellValue(this.Point);
        set
        {
            this._cellsCollection.ValueSlice.SetCellValue(this.Point, value);
            this.Worksheet.Workbook.CalcEngine.MarkDirty(this.Worksheet, this.Point);
        }
    }

    private XLImmutableRichText SliceRichText
    {
        get => this._cellsCollection.ValueSlice.GetRichText(this.Point);
        set => this._cellsCollection.ValueSlice.SetRichText(this.Point, value);
    }

    private XLComment SliceComment
    {
        get => this._cellsCollection.MiscSlice[this._rowNumber, this._columnNumber].Comment;
        set
        {
            ref readonly XLMiscSliceContent original = ref this._cellsCollection.MiscSlice[
                this._rowNumber,
                this._columnNumber
            ];
            if (original.Comment != value)
            {
                XLMiscSliceContent modified = original;
                modified.Comment = value;
                this._cellsCollection.MiscSlice.Set(
                    this._rowNumber,
                    this._columnNumber,
                    in modified
                );
            }
        }
    }

    internal uint? CellMetaIndex
    {
        get => this._cellsCollection.MiscSlice[this._rowNumber, this._columnNumber].CellMetaIndex;
        set
        {
            ref readonly XLMiscSliceContent original = ref this._cellsCollection.MiscSlice[
                this._rowNumber,
                this._columnNumber
            ];
            if (original.CellMetaIndex != value)
            {
                XLMiscSliceContent modified = original;
                modified.CellMetaIndex = value;
                this._cellsCollection.MiscSlice.Set(
                    this._rowNumber,
                    this._columnNumber,
                    in modified
                );
            }
        }
    }

    internal uint? ValueMetaIndex
    {
        get => this._cellsCollection.MiscSlice[this._rowNumber, this._columnNumber].ValueMetaIndex;
        set
        {
            ref readonly XLMiscSliceContent original = ref this._cellsCollection.MiscSlice[
                this._rowNumber,
                this._columnNumber
            ];
            if (original.ValueMetaIndex != value)
            {
                XLMiscSliceContent modified = original;
                modified.ValueMetaIndex = value;
                this._cellsCollection.MiscSlice.Set(
                    this._rowNumber,
                    this._columnNumber,
                    in modified
                );
            }
        }
    }

    /// <summary>
    /// A formula in the cell. Null, if cell doesn't contain formula.
    /// </summary>
    internal XLCellFormula Formula
    {
        get => this._cellsCollection.FormulaSlice.Get(this.Point);
        set
        {
            this._cellsCollection.FormulaSlice.Set(this.Point, value);

            // Because text values of evaluated formulas are stored in a worksheet part, mark it as inlined string and store in sst.
            // If we are clearing formula, we should enable shareString back on, because it is a default position.
            // If we are setting formula, we should disable shareString (=inline), because it must be written to the worksheet part
            bool clearFormula = value is null;
            this.ShareString = clearFormula;
            this.Worksheet.Workbook.CalcEngine.MarkDirty(this.Worksheet, this.Point);
        }
    }

    #endregion Slice fields

    #region IXLFormatContainer
#nullable enable

    /// <summary>
    /// A format of a cell. If cell format depends on inherited format, the value is <c>null</c>.
    /// </summary>
    public XLCellFormatValue? FormatValue
    {
        get => this._cellsCollection.FormatSlice.GetFormat(this.Point);
        set => this._cellsCollection.FormatSlice.Set(this.Point, value);
    }
#nullable disable
    #endregion

    internal XLCellFormat Format => XLCellFormat.ForCell(this);

    internal XLComment GetComment() => this.SliceComment ?? this.CreateComment();

    internal XLComment CreateComment(int? shapeId = null) =>
        this.SliceComment = XLComment.Create(this, shapeId: shapeId);

    public XLRichText GetRichText()
    {
        XLImmutableRichText sliceRichText = this.SliceRichText;
        if (sliceRichText is not null)
        {
            XLCellFormatValue cellFormat = this.GetFormat();
            return new XLRichText(this, cellFormat.Font, sliceRichText);
        }

        return this.CreateRichText();
    }

    public XLRichText CreateRichText()
    {
        XLFontFormatValue fontFormat = this.GetFormat().Font;

        // Don't include rich text string with 0 length to a new rich text
        XLRichText richText =
            this.DataType == XLDataType.Blank
                ? new XLRichText(this, fontFormat)
                : new XLRichText(this, fontFormat, this.GetFormattedString());
        this.SliceRichText = XLImmutableRichText.Create(richText);
        return richText;
    }

    #region IXLCell Members

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    IXLWorksheet IXLCell.Worksheet => this.Worksheet;

    IXLAddress IXLCell.Address => this.Address;

    IXLRange IXLCell.AsRange() => this.AsRange();

    internal IXLCell SetValue(XLCellValue value, bool setTableHeader, bool checkMergedRanges)
    {
        if (checkMergedRanges && this.IsInferiorMergedCell())
        {
            return this;
        }

        // Mimic Excel behavior: When a value is set to a a certain types (e.g. timespan or
        // a date), the format of a cell is changed.
        XLCellFormatValue valueRequiredFormat = this.Worksheet.GetStyleForValue(value, this.Point);
        if (valueRequiredFormat is not null)
        {
            this.FormatValue = valueRequiredFormat;
        }

        // Modify value after style, because we might need to strip the '
        if (value.Type == XLDataType.Text)
        {
            string text = value.GetText();
            if (text.Length > 0 && text[0] == '\'')
            {
                value = text.Substring(1);
            }
        }

        this.SetOnlyValue(value);

        this.FormulaA1 = null;

        if (setTableHeader)
        {
            Area cellRange = new(this.Point, this.Point);
            foreach (XLTable table in this.Worksheet.Tables)
            {
                table.RefreshFieldsFromCells(cellRange);
            }
        }

        return this;
    }

    public bool GetBoolean() => this.Value.GetBoolean();

    public double GetDouble() => this.Value.GetNumber();

    public string GetText() => this.Value.GetText();

    public XLError GetError() => this.Value.GetError();

    public DateTime GetDateTime() => this.Value.GetDateTime();

    public TimeSpan GetTimeSpan() => this.Value.GetTimeSpan();

    public bool TryGetValue<T>(out T value)
    {
        XLCellValue currentValue;
        try
        {
            currentValue = this.Value;
        }
        catch
        {
            // May fail for formula evaluation
            value = default;
            return false;
        }

        Type targetType = typeof(T);
        bool isNullable = targetType.IsNullableType();
        if (isNullable && currentValue.TryConvert(out Blank _))
        {
            value = default;
            return true;
        }

        // JIT compiles a separate version for each T value type and one for all reference types
        // Optimization then removes the double casting for value types.
        Type underlyingType = targetType.GetUnderlyingType();
        if (underlyingType == typeof(DateTime) && currentValue.TryConvert(out DateTime dateTime))
        {
            value = (T)(object)dateTime;
            return true;
        }

        CultureInfo culture = CultureInfo.CurrentCulture;
        if (
            underlyingType == typeof(TimeSpan)
            && currentValue.TryConvert(out TimeSpan timeSpan, culture)
        )
        {
            value = (T)(object)timeSpan;
            return true;
        }

        if (underlyingType == typeof(bool) && currentValue.TryConvert(out bool boolean))
        {
            value = (T)(object)boolean;
            return true;
        }

        if (TryGetStringValue(out value, currentValue))
        {
            return true;
        }

        if (underlyingType == typeof(XLError))
        {
            if (currentValue.IsError)
            {
                value = (T)(object)currentValue.GetError();
                return true;
            }

            return false;
        }

        // Type code of an enum is a type of an integer, so do this check before numbers
        if (underlyingType.IsEnum)
        {
            string strValue = currentValue.ToString(culture);
            if (Enum.IsDefined(underlyingType, strValue))
            {
                value = (T)Enum.Parse(underlyingType, strValue, ignoreCase: false);
                return true;
            }
            value = default;
            return false;
        }

        TypeCode typeCode = Type.GetTypeCode(underlyingType);

        // T is a floating point numbers
        if (typeCode >= TypeCode.Single && typeCode <= TypeCode.Decimal)
        {
            if (!currentValue.TryConvert(out double doubleValue, culture))
            {
                return false;
            }

            if (typeCode == TypeCode.Single && doubleValue is < float.MinValue or > float.MaxValue)
            {
                return false;
            }

            value = typeCode switch
            {
                TypeCode.Single => (T)(object)(float)doubleValue,
                TypeCode.Double => (T)(object)doubleValue,
                TypeCode.Decimal => (T)(object)(decimal)doubleValue,
                _ => throw new NotSupportedException(),
            };
            return true;
        }

        // T is an integer
        if (typeCode >= TypeCode.SByte && typeCode <= TypeCode.UInt64)
        {
            if (!currentValue.TryConvert(out double doubleValue, culture))
            {
                return false;
            }

            if (!doubleValue.Equals(Math.Truncate(doubleValue)))
            {
                return false;
            }

            bool valueIsWithinBounds = typeCode switch
            {
                TypeCode.SByte => doubleValue >= sbyte.MinValue && doubleValue <= sbyte.MaxValue,
                TypeCode.Byte => doubleValue >= byte.MinValue && doubleValue <= byte.MaxValue,
                TypeCode.Int16 => doubleValue >= short.MinValue && doubleValue <= short.MaxValue,
                TypeCode.UInt16 => doubleValue >= ushort.MinValue && doubleValue <= ushort.MaxValue,
                TypeCode.Int32 => doubleValue >= int.MinValue && doubleValue <= int.MaxValue,
                TypeCode.UInt32 => doubleValue >= uint.MinValue && doubleValue <= uint.MaxValue,
                TypeCode.Int64 => doubleValue >= long.MinValue && doubleValue <= long.MaxValue,
                TypeCode.UInt64 => doubleValue >= ulong.MinValue && doubleValue <= ulong.MaxValue,
                _ => throw new NotSupportedException(),
            };
            if (!valueIsWithinBounds)
            {
                return false;
            }

            value = typeCode switch
            {
                TypeCode.SByte => (T)(object)(sbyte)doubleValue,
                TypeCode.Byte => (T)(object)(byte)doubleValue,
                TypeCode.Int16 => (T)(object)(short)doubleValue,
                TypeCode.UInt16 => (T)(object)(ushort)doubleValue,
                TypeCode.Int32 => (T)(object)(int)doubleValue,
                TypeCode.UInt32 => (T)(object)(uint)doubleValue,
                TypeCode.Int64 => (T)(object)(long)doubleValue,
                TypeCode.UInt64 => (T)(object)(ulong)doubleValue,
                _ => throw new NotSupportedException(),
            };
            return true;
        }

        return false;
    }

    private static bool TryGetStringValue<T>(out T value, XLCellValue currentValue)
    {
        if (typeof(T) == typeof(string))
        {
            string s = currentValue.ToString(CultureInfo.CurrentCulture);
            MatchCollection matches = utfPattern.Matches(s);

            if (matches.Count == 0)
            {
                value = (T)Convert.ChangeType(s, typeof(T));
                return true;
            }

            StringBuilder sb = new();
            int lastIndex = 0;

            foreach (Match match in matches.Cast<Match>())
            {
                string matchString = match.Value;
                int matchIndex = match.Index;
                sb.Append(s.Substring(lastIndex, matchIndex - lastIndex));

                sb.Append((char)int.Parse(match.Groups[1].Value, NumberStyles.AllowHexSpecifier));

                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < s.Length)
            {
                sb.Append(s.Substring(lastIndex));
            }

            value = (T)Convert.ChangeType(sb.ToString(), typeof(T));
            return true;
        }
        value = default;
        return false;
    }

    public T GetValue<T>()
    {
        if (this.TryGetValue(out T retVal))
        {
            return retVal;
        }

        throw new InvalidCastException(
            $"Cannot convert {this.Address.ToStringRelative(true)}'s value to " + typeof(T)
        );
    }

    public string GetString() => this.Value.ToString(CultureInfo.CurrentCulture);

    public string GetFormattedString(CultureInfo culture = null)
    {
        XLCellValue value;
        try
        {
            // Need to get actual value because formula might be out of date or value wasn't set at all
            // Unimplemented functions and features throw exceptions
            value = this.Value;
        }
        catch
        {
            value = this.CachedValue;
        }

        return this.GetFormattedString(value, culture);
    }

    internal string GetFormattedString(XLCellValue value, CultureInfo culture = null)
    {
        culture ??= CultureInfo.CurrentCulture;
        XLNumberFormat numberFormat = this.GetFormat().NumberFormat;
        return value.IsUnifiedNumber
            ? value.GetUnifiedNumber().ToExcelFormat(numberFormat, culture)
            : value.ToString(culture);
    }

    public void InvalidateFormula()
    {
        if (this.Formula is null)
        {
            return;
        }

        this.Formula.IsDirty = true;
    }

    /// <summary>
    /// Perform an evaluation of cell formula. If cell does not contain formula nothing happens, if cell does not need
    /// recalculation (<see cref="NeedsRecalculation"/> is False) nothing happens either, unless <paramref name="force"/> flag is specified.
    /// Otherwise recalculation is performed, result value is preserved in <see cref="CachedValue"/> and returned.
    /// </summary>
    /// <param name="force">Flag indicating whether a recalculation must be performed even is cell does not need it.</param>
    /// <returns>Null if cell does not contain a formula. Calculated value otherwise.</returns>
    public void Evaluate(bool force)
    {
        if (this.Formula is null)
        {
            return;
        }

        bool shouldRecalculate = force || this.NeedsRecalculation;
        if (!shouldRecalculate)
        {
            return;
        }

        // TODO: Only one cell, somehow
        XLWorkbook wb = this.Worksheet.Workbook;
        wb.CalcEngine.Recalculate(wb, null);
    }

    /// <summary>
    /// Set only value, don't clear formula, don't set format.
    /// Sets the value even for merged cells.
    /// </summary>
    internal void SetOnlyValue(XLCellValue value) => this.SliceCellValue = value;

    public IXLCell SetValue(XLCellValue value) => this.SetValue(value, true, true);

    public override string ToString() => this.ToString("A");

    public string ToString(string format) =>
        (format.ToUpper()) switch
        {
            "A" => this.Address.ToString(),
            "F" => this.HasFormula ? this.FormulaA1 : string.Empty,
            "NF" => this.Style.NumberFormat.Format,
            "FG" => this.Style.Font.FontColor.ToString(),
            "BG" => this.Style.Fill.BackgroundColor.ToString(),
            "V" => this.GetFormattedString(),
            _ => throw new FormatException($"Format {format} was not recognised."),
        };

    public XLCellValue Value
    {
        get
        {
            if (this.Formula is not null)
            {
                this.Evaluate(false);
            }

            return this.SliceCellValue;
        }
        set => this.SetValue(value);
    }

    public IXLTable InsertTable<T>(IEnumerable<T> data) => this.InsertTable(data, null, true);

    public IXLTable InsertTable<T>(IEnumerable<T> data, bool createTable) =>
        this.InsertTable(data, null, createTable);

    public IXLTable InsertTable<T>(IEnumerable<T> data, string tableName) =>
        this.InsertTable(data, tableName, true);

    public IXLTable InsertTable<T>(IEnumerable<T> data, string tableName, bool createTable) =>
        this.InsertTable(data, tableName, createTable, addHeadings: true, transpose: false);

    public IXLTable InsertTable<T>(
        IEnumerable<T> data,
        string tableName,
        bool createTable,
        bool addHeadings,
        bool transpose
    )
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        return this.Worksheet.InsertTable(
            this.Point,
            reader,
            tableName,
            createTable,
            addHeadings,
            transpose
        );
    }

    public IXLTable InsertTable(DataTable data) => this.InsertTable(data, null, true);

    public IXLTable InsertTable(DataTable data, bool createTable) =>
        this.InsertTable(data, null, createTable);

    public IXLTable InsertTable(DataTable data, string tableName) =>
        this.InsertTable(data, tableName, true);

    public IXLTable InsertTable(DataTable data, string tableName, bool createTable)
    {
        if (data == null || data.Columns.Count == 0)
        {
            return null;
        }

        if (
            XlsxSharp.XLHelper.IsValidA1Address(tableName)
            || XlsxSharp.XLHelper.IsValidRCAddress(tableName)
        )
        {
            throw new InvalidOperationException(
                $"Table name cannot be a valid Cell Address '{tableName}'."
            );
        }

        if (createTable && this.Worksheet.Tables.Any<XLTable>(t => t.Contains(this)))
        {
            throw new InvalidOperationException(
                $"This cell '{this.Address}' is already part of a table."
            );
        }

        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        return this.Worksheet.InsertTable(
            this.Point,
            reader,
            tableName,
            createTable,
            addHeadings: true,
            transpose: false
        );
    }

    public XLTableCellType TableCellType()
    {
        XLTable table = this.Worksheet.Tables.FirstOrDefault<XLTable>(t =>
            t.AsRange().Contains(this)
        );
        if (table == null)
        {
            return XLTableCellType.None;
        }

        if (table.ShowHeaderRow && table.HeadersRow().RowNumber().Equals(this._rowNumber))
        {
            return XLTableCellType.Header;
        }

        if (table.ShowTotalsRow && table.TotalsRow().RowNumber().Equals(this._rowNumber))
        {
            return XLTableCellType.Total;
        }

        return XLTableCellType.Data;
    }

    public IXLRange InsertData(IEnumerable data)
    {
        if (data == null || data is string)
        {
            return null;
        }

        return this.InsertData(data, transpose: false);
    }

    public IXLRange InsertData(IEnumerable data, bool transpose)
    {
        if (data == null || data is string)
        {
            return null;
        }

        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        return this.Worksheet.InsertData(
            this.Point,
            reader,
            addHeadings: false,
            transpose: transpose
        );
    }

    public IXLRange InsertData(DataTable dataTable)
    {
        if (dataTable == null)
        {
            return null;
        }

        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(dataTable);
        return this.Worksheet.InsertData(this.Point, reader, addHeadings: false, transpose: false);
    }

    public XLDataType DataType => this.SliceCellValue.Type;

    public IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.All) =>
        this.Clear(clearOptions, false);

    internal IXLCell Clear(XLClearOptions clearOptions, bool calledFromRange)
    {
        //Note: We have to check if the cell is part of a merged range. If so we have to clear the whole range
        //Checking if called from range to avoid stack overflow
        if (!calledFromRange && this.IsMerged())
        {
            IXLRange firstOrDefault = this
                .Worksheet.Internals.MergedRanges.GetIntersectedRanges(this.Address)
                .FirstOrDefault();
            if (firstOrDefault != null)
            {
                firstOrDefault.Clear(clearOptions);
            }
        }
        else
        {
            if (clearOptions.HasFlag(XLClearOptions.Contents))
            {
                this.SetHyperlink(null);
                this.SliceCellValue = Blank.Value;
                this.FormulaA1 = string.Empty;
            }

            if (clearOptions.HasFlag(XLClearOptions.NormalFormats))
            {
                this.FormatValue = this.Worksheet.FormatValue;
            }

            if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
            {
                this.Worksheet.ConditionalFormats.Clear(this.Point);
            }

            if (clearOptions.HasFlag(XLClearOptions.Comments))
            {
                this.SliceComment = null;
            }

            if (clearOptions.HasFlag(XLClearOptions.Sparklines))
            {
                this.AsRange().RemoveSparklines();
            }

            if (clearOptions.HasFlag(XLClearOptions.DataValidation) && this.HasDataValidation)
            {
                XLDataValidation validation = this.CreateDataValidation();
                this.Worksheet.DataValidations.Delete(validation);
            }

            if (clearOptions.HasFlag(XLClearOptions.MergedRanges) && this.IsMerged())
            {
                this.ClearMerged();
            }
        }

        return this;
    }

    public void Delete(XLShiftDeletedCells shiftDeleteCells) =>
        this.Worksheet.Range(this.Address, this.Address).Delete(shiftDeleteCells);

    public string FormulaA1
    {
        get => this.Formula?.A1 ?? string.Empty;
        set
        {
            if (this.IsInferiorMergedCell())
            {
                return;
            }

            string formula = value?.TrimFormulaEqual();
            if (!string.IsNullOrWhiteSpace(formula))
            {
                string fixedFunctionsFormula = FormulaTransformation.FixFutureFunctions(
                    formula,
                    this.Worksheet.Name,
                    this.Point
                );
                this.Formula = XLCellFormula.NormalA1(fixedFunctionsFormula);
            }
            else
            {
                this.Formula = null;
            }

            this.InvalidateFormula();
        }
    }

    public string FormulaR1C1
    {
        get => this.Formula?.GetFormulaR1C1(this.Point) ?? string.Empty;
        set
        {
            if (this.IsInferiorMergedCell())
            {
                return;
            }

            string formula = value?.TrimFormulaEqual();
            if (!string.IsNullOrWhiteSpace(formula))
            {
                string formulaA1 = FormulaConverter.ToA1(
                    formula,
                    this._rowNumber,
                    this._columnNumber
                );
                string fixedFunctionsFormulaA1 = FormulaTransformation.FixFutureFunctions(
                    formulaA1,
                    this.Worksheet.Name,
                    this.Point
                );
                this.Formula = XLCellFormula.NormalA1(fixedFunctionsFormulaA1);
            }
            else
            {
                this.Formula = null;
            }

            this.InvalidateFormula();
        }
    }

    public XLHyperlink GetHyperlink()
    {
        if (this.Worksheet.Hyperlinks.TryGet(this.Point, out XLHyperlink hyperlink))
        {
            return hyperlink;
        }

        return this.CreateHyperlink();
    }

#nullable enable
    /// <inheritdoc />
    public void SetHyperlink(XLHyperlink? hyperlink)
    {
        this.Worksheet.Hyperlinks.SetCellHyperlink(this.Point, hyperlink);
        if (hyperlink is null)
        {
            return;
        }

        XLCellFormatValue cellFormat = this.GetFormat();
        XLCellFormatValue sheetFormat = this.Worksheet.GetFormat();
        if (ReferenceEquals(cellFormat, sheetFormat))
        {
            this.FormatValue = this.Styles.GetModifiedFormat(
                cellFormat,
                font =>
                    font with
                    {
                        Color = XLColor.FromTheme(XLThemeColor.Hyperlink),
                        Underline = XLFontUnderlineValues.Single,
                    }
            );
        }
    }

    internal void SetCellHyperlink(XLHyperlink hyperlink) =>
        this.Worksheet.Hyperlinks.SetCellHyperlink(this.Point, hyperlink);

#nullable disable

    public XLHyperlink CreateHyperlink()
    {
        XLHyperlink link = new();
        this.SetHyperlink(link);
        return link;
    }

    public IXLCells InsertCellsAbove(int numberOfRows) =>
        this.AsRange().InsertRowsAbove(numberOfRows).Cells();

    public IXLCells InsertCellsBelow(int numberOfRows) =>
        this.AsRange().InsertRowsBelow(numberOfRows).Cells();

    public IXLCells InsertCellsAfter(int numberOfColumns) =>
        this.AsRange().InsertColumnsAfter(numberOfColumns).Cells();

    public IXLCells InsertCellsBefore(int numberOfColumns) =>
        this.AsRange().InsertColumnsBefore(numberOfColumns).Cells();

    public IXLCell AddToNamed(string rangeName)
    {
        this.AsRange().AddToNamed(rangeName);
        return this;
    }

    public IXLCell AddToNamed(string rangeName, XLScope scope)
    {
        this.AsRange().AddToNamed(rangeName, scope);
        return this;
    }

    public IXLCell AddToNamed(string rangeName, XLScope scope, string comment)
    {
        this.AsRange().AddToNamed(rangeName, scope, comment);
        return this;
    }

    /// <summary>
    /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
    /// </summary>
    public bool NeedsRecalculation => this.Formula is not null && this.Formula.IsDirty;

    public XLCellValue CachedValue => this.SliceCellValue;

    IXLRichText IXLCell.GetRichText() => this.GetRichText();

    public bool HasRichText => this.SliceRichText is not null;

    IXLRichText IXLCell.CreateRichText() => this.CreateRichText();

    IXLComment IXLCell.GetComment() => this.GetComment();

    public bool HasComment => this.SliceComment != null;

    IXLComment IXLCell.CreateComment() => this.CreateComment(shapeId: null);

    public bool IsMerged() => this.Worksheet.Internals.MergedRanges.Contains(this);

    public IXLRange MergedRange() =>
        this.Worksheet.Internals.MergedRanges.GetIntersectedRanges(this).FirstOrDefault();

    public bool IsEmpty() => this.IsEmpty(XLCellsUsedOptions.AllContents);

    public bool IsEmpty(XLCellsUsedOptions options)
    {
        bool isValueEmpty = this.SliceCellValue.Type switch
        {
            XLDataType.Blank => true,
            XLDataType.Text => this.SliceCellValue.GetText().Length == 0,
            _ => false,
        };

        if (!isValueEmpty || this.HasFormula)
        {
            return false;
        }

        if (options.HasFlag(XLCellsUsedOptions.NormalFormats))
        {
            if (this.FormatValue is { } cellFormat)
            {
                XLCellFormatValue inheritedFormat = this.Worksheet.GetInheritedFormat(this.Point);
                if (!XLCellFormatValue.AreSame(cellFormat, inheritedFormat))
                {
                    return false;
                }
            }
        }

        if (options.HasFlag(XLCellsUsedOptions.MergedRanges) && this.IsMerged())
        {
            return false;
        }

        if (options.HasFlag(XLCellsUsedOptions.Comments) && this.HasComment)
        {
            return false;
        }

        if (options.HasFlag(XLCellsUsedOptions.DataValidation) && this.HasDataValidation)
        {
            return false;
        }

        if (
            options.HasFlag(XLCellsUsedOptions.ConditionalFormats)
            && this.Worksheet.ConditionalFormats.SelectMany<XLConditionalFormat, IXLRange>(cf =>
                    cf.Ranges
                )
                .Any(range => range.Contains(this))
        )
        {
            return false;
        }

        if (options.HasFlag(XLCellsUsedOptions.Sparklines) && this.HasSparkline)
        {
            return false;
        }

        return true;
    }

    public IXLColumn WorksheetColumn() => this.Worksheet.Column(this._columnNumber);

    public IXLRow WorksheetRow() => this.Worksheet.Row(this._rowNumber);

    public IXLCell CopyTo(IXLCell target)
    {
        (target as XLCell).CopyFrom(this, XLCellCopyOptions.All);
        return target;
    }

    public IXLCell CopyTo(string target) => this.CopyTo(GetTargetCell(target, this.Worksheet));

    public IXLCell CopyFrom(IXLCell otherCell) =>
        this.CopyFrom(otherCell as XLCell, XLCellCopyOptions.All);

    public IXLCell CopyFrom(string otherCell) =>
        this.CopyFrom(GetTargetCell(otherCell, this.Worksheet));

    public IXLCell SetFormulaA1(string formula)
    {
        this.FormulaA1 = formula;
        return this;
    }

    public IXLCell SetFormulaR1C1(string formula)
    {
        this.FormulaR1C1 = formula;
        return this;
    }

    public bool HasSparkline => this.Sparkline != null;

    /// <summary> The sparkline assigned to the cell </summary>
    public IXLSparkline Sparkline => this.Worksheet.SparklineGroups.GetSparkline(this);

    public IXLDataValidation GetDataValidation() =>
        this.FindDataValidation() ?? this.CreateDataValidation();

    public bool HasDataValidation => this.FindDataValidation() != null;

    /// <summary>
    /// Get the data validation rule containing current cell.
    /// </summary>
    /// <returns>The data validation rule applying to the current cell or null if there is no such rule.</returns>
    private IXLDataValidation FindDataValidation()
    {
        this.Worksheet.DataValidations.TryGet(
            new XLRangeAddress(this.Address, this.Address),
            out IXLDataValidation dataValidation
        );
        return dataValidation;
    }

    IXLDataValidation IXLCell.CreateDataValidation() => this.CreateDataValidation();

    internal XLDataValidation CreateDataValidation() =>
        this.Worksheet.DataValidations.Create(new Area(this.Point));

    public void Select() => this.AsRange().Select();

    public IXLConditionalFormat AddConditionalFormat() => this.AsRange().AddConditionalFormat();

    public bool Active
    {
        get => this.Worksheet.ActiveCell == this.Point;
        set
        {
            if (value)
            {
                this.Worksheet.ActiveCell = this.Point;
            }
            else if (this.Active)
            {
                this.Worksheet.ActiveCell = null;
            }
        }
    }

    public IXLCell SetActive(bool value = true)
    {
        this.Active = value;
        return this;
    }

    public bool HasHyperlink => this.Worksheet.Hyperlinks.HasHyperlink(this.Point);

    /// <inheritdoc />
    public bool ShowPhonetic
    {
        get => this._cellsCollection.MiscSlice[this._rowNumber, this._columnNumber].HasPhonetic;
        set
        {
            ref readonly XLMiscSliceContent original = ref this._cellsCollection.MiscSlice[
                this._rowNumber,
                this._columnNumber
            ];
            if (original.HasPhonetic != value)
            {
                XLMiscSliceContent modified = original;
                modified.HasPhonetic = value;
                this._cellsCollection.MiscSlice.Set(
                    this._rowNumber,
                    this._columnNumber,
                    in modified
                );
            }
        }
    }

    #endregion IXLCell Members

    /// <summary>
    /// Ensure the cell has style set directly on the cell, not inherited from column/row/worksheet styles.
    /// </summary>
    internal void PingStyle() => this.FormatValue = this.GetFormat();

    public XLRange AsRange() => this.Worksheet.Range(this.Address, this.Address);

    #region Styles

    /// <summary>
    /// Get format of a cell that should be used to render it.
    /// </summary>
    internal XLCellFormatValue GetFormat() => this.Worksheet.GetStyleValue(this.Point);

    #endregion Styles

    public void DeleteComment() => this.Clear(XLClearOptions.Comments);

    public void DeleteSparkline() => this.Clear(XLClearOptions.Sparklines);

    public IXLCell CopyFrom(IXLRangeBase rangeObject)
    {
        ArgumentNullException.ThrowIfNull(rangeObject);

        XLRangeBase asRange = (XLRangeBase)rangeObject;
        int maxRows = asRange.RowCount();
        int maxColumns = asRange.ColumnCount();

        int lastRow = Math.Min(this._rowNumber + maxRows - 1, XlsxSharp.XLHelper.MaxRowNumber);
        int lastColumn = Math.Min(
            this._columnNumber + maxColumns - 1,
            XlsxSharp.XLHelper.MaxColumnNumber
        );

        XLRange targetRange = this.Worksheet.Range(
            this._rowNumber,
            this._columnNumber,
            lastRow,
            lastColumn
        );

        if (!(asRange is XLRow || asRange is XLColumn))
        {
            targetRange.Clear();
        }

        int minRow = asRange.RangeAddress.FirstAddress.RowNumber;
        int minColumn = asRange.RangeAddress.FirstAddress.ColumnNumber;
        IXLCells cellsUsed = asRange.CellsUsed(
            XLCellsUsedOptions.All
                & ~XLCellsUsedOptions.ConditionalFormats
                & ~XLCellsUsedOptions.DataValidation
                & ~XLCellsUsedOptions.MergedRanges
        );
        foreach (IXLCell sourceCell in cellsUsed)
        {
            this.Worksheet.Cell(
                    this._rowNumber + sourceCell.Address.RowNumber - minRow,
                    this._columnNumber + sourceCell.Address.ColumnNumber - minColumn
                )
                .CopyFromInternal(
                    sourceCell as XLCell,
                    XLCellCopyOptions.All
                        & ~XLCellCopyOptions.ConditionalFormats
                        & ~XLCellCopyOptions.DataValidations
                ); //Conditional formats and data validation are copied separately
        }

        List<IXLRange> rangesToMerge =
        [
            .. asRange
                .Worksheet.Internals.MergedRanges.Where<XLRange>(mr => asRange.Contains(mr))
                .Select(mr =>
                {
                    int firstRow =
                        this._rowNumber
                        + (
                            mr.RangeAddress.FirstAddress.RowNumber
                            - asRange.RangeAddress.FirstAddress.RowNumber
                        );
                    int firstColumn =
                        this._columnNumber
                        + (
                            mr.RangeAddress.FirstAddress.ColumnNumber
                            - asRange.RangeAddress.FirstAddress.ColumnNumber
                        );
                    return (IXLRange)
                        this.Worksheet.Range(
                            firstRow,
                            firstColumn,
                            firstRow + mr.RowCount() - 1,
                            firstColumn + mr.ColumnCount() - 1
                        );
                }),
        ];

        rangesToMerge.ForEach(r => r.Merge(false));

        List<IXLDataValidation> dataValidations =
        [
            .. asRange.Worksheet.DataValidations.GetAllInRange(asRange.RangeAddress),
        ];

        foreach (IXLDataValidation dataValidation in dataValidations)
        {
            XLDataValidation newDataValidation = null;
            foreach (IXLRange dvRange in dataValidation.Ranges.Where(r => r.Intersects(asRange)))
            {
                IXLRangeAddress dvTargetAddress = dvRange.RangeAddress.Relative(
                    asRange.RangeAddress,
                    targetRange.RangeAddress
                );
                XLRange dvTargetRange = this.Worksheet.Range(dvTargetAddress);
                if (newDataValidation == null)
                {
                    newDataValidation = dvTargetRange.CreateDataValidation();
                    newDataValidation.CopyFrom(dataValidation);
                }
                else
                {
                    newDataValidation.AddRange(dvTargetRange);
                }
            }
        }

        this.Worksheet.ConditionalFormats.CopyFrom(
            asRange.Worksheet,
            asRange.SheetRange,
            this.Point
        );
        return this;
    }

    private void ClearMerged()
    {
        List<IXLRange> mergeToDelete =
        [
            .. this.Worksheet.Internals.MergedRanges.GetIntersectedRanges(this.Address),
        ];

        mergeToDelete.ForEach(m => this.Worksheet.Internals.MergedRanges.Remove(m));
    }

    internal string GetFormulaR1C1(string value) =>
        XLCellFormula.GetFormula(
            value,
            FormulaConversionType.A1ToR1C1,
            new Point(this._rowNumber, this._columnNumber)
        );

    internal string GetFormulaA1(string value) =>
        XLCellFormula.GetFormula(
            value,
            FormulaConversionType.R1C1ToA1,
            new Point(this._rowNumber, this._columnNumber)
        );

    internal void CopyValuesFrom(XLCell source)
    {
        // Rich text is basically a super set of a value. Setting a value would override rich text and vice versa.
        XLImmutableRichText sourceRichText = source.SliceRichText;
        if (sourceRichText is null)
        {
            this.SliceCellValue = source.SliceCellValue;
        }
        else
        {
            this.SliceRichText = sourceRichText;
        }

        this.FormulaR1C1 = source.FormulaR1C1;
        this.SliceComment =
            source.SliceComment == null
                ? null
                : XLComment.CreateAsCopy(this, source, source.SliceComment);

        if (source.Worksheet.Hyperlinks.TryGet(source.Point, out XLHyperlink sourceHyperlink))
        {
            this.SetCellHyperlink(new XLHyperlink(sourceHyperlink));
        }
    }

    private static IXLCell GetTargetCell(string target, XLWorksheet defaultWorksheet)
    {
        string[] pair = target.Split('!');
        if (pair.Length == 1)
        {
            return defaultWorksheet.Cell(target);
        }

        string wsName = pair[0];
        if (wsName.StartsWith('\''))
        {
            wsName = wsName.Substring(1, wsName.Length - 2);
        }

        return defaultWorksheet.Workbook.Worksheet(wsName).Cell(pair[1]);
    }

    internal IXLCell CopyFromInternal(XLCell otherCell, XLCellCopyOptions options)
    {
        if (options.HasFlag(XLCellCopyOptions.Values))
        {
            this.CopyValuesFrom(otherCell);
        }

        // Other cell might be from a different workbook.
        if (options.HasFlag(XLCellCopyOptions.Styles))
        {
            this.FormatValue = this.Styles.GetRegisteredCellFormat(otherCell.GetFormat());
        }

        if (options.HasFlag(XLCellCopyOptions.Sparklines))
        {
            this.CopySparklineFrom(otherCell);
        }

        if (options.HasFlag(XLCellCopyOptions.ConditionalFormats))
        {
            this.Worksheet.ConditionalFormats.CopyFrom(
                otherCell.Worksheet,
                otherCell.Point,
                this.Point,
                true
            );
        }

        if (options.HasFlag(XLCellCopyOptions.DataValidations))
        {
            this.CopyDataValidationFrom(otherCell);
        }

        return this;
    }

    private void CopySparklineFrom(XLCell otherCell)
    {
        if (!otherCell.HasSparkline)
        {
            return;
        }

        string sourceDataAddress = otherCell.Sparkline.SourceData.RangeAddress.ToString();
        string shiftedRangeAddress = this.GetFormulaA1(otherCell.GetFormulaR1C1(sourceDataAddress));
        XLWorksheet sourceDataWorksheet =
            otherCell.Worksheet == otherCell.Sparkline.SourceData.Worksheet
                ? this.Worksheet
                : (XLWorksheet)otherCell.Sparkline.SourceData.Worksheet;
        XLRange sourceData = sourceDataWorksheet.Range(shiftedRangeAddress);

        IXLSparklineGroup group;
        if (otherCell.Worksheet == this.Worksheet)
        {
            group = otherCell.Sparkline.SparklineGroup;
        }
        else
        {
            group = this.Worksheet.SparklineGroups.Add(
                new XLSparklineGroup(this.Worksheet, otherCell.Sparkline.SparklineGroup)
            );
            if (otherCell.Sparkline.SparklineGroup.DateRange != null)
            {
                IXLWorksheet dateRangeWorksheet =
                    otherCell.Worksheet == otherCell.Sparkline.SparklineGroup.DateRange.Worksheet
                        ? this.Worksheet
                        : otherCell.Sparkline.SparklineGroup.DateRange.Worksheet;
                string dateRangeAddress =
                    otherCell.Sparkline.SparklineGroup.DateRange.RangeAddress.ToString();
                string shiftedDateRangeAddress = this.GetFormulaA1(
                    otherCell.GetFormulaR1C1(dateRangeAddress)
                );
                group.SetDateRange(dateRangeWorksheet.Range(shiftedDateRangeAddress));
            }
        }

        group.Add(this, sourceData);
    }

    public IXLCell CopyFrom(IXLCell otherCell, XLCellCopyOptions options)
    {
        XLCell source = otherCell as XLCell; // To expose GetFormulaR1C1, etc

        this.CopyFromInternal(source, options);
        return this;
    }

    private void CopyDataValidationFrom(XLCell otherCell)
    {
        if (otherCell.HasDataValidation)
        {
            this.CopyDataValidation(otherCell, otherCell.GetDataValidation());
        }
        else if (this.HasDataValidation)
        {
            this.Worksheet.DataValidations.Delete(new Area(this.Point));
        }
    }

    internal void CopyDataValidation(XLCell otherCell, IXLDataValidation otherDv)
    {
        XLDataValidation thisDv = this.GetDataValidation() as XLDataValidation;
        thisDv.CopyFrom(otherDv);
        thisDv.Value = this.GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.Value));
        thisDv.MinValue = this.GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MinValue));
        thisDv.MaxValue = this.GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MaxValue));
    }

    internal void ShiftFormulaRows(XLRange shiftedRange, int rowsShifted) =>
        this.FormulaA1 = ShiftFormulaRows(
            this.FormulaA1,
            this.Worksheet,
            shiftedRange,
            rowsShifted
        );

    internal static string ShiftFormulaRows(
        string formulaA1,
        XLWorksheet worksheetInAction,
        XLRange shiftedRange,
        int rowsShifted
    )
    {
        if (string.IsNullOrWhiteSpace(formulaA1))
        {
            return string.Empty;
        }

        string value = formulaA1;

        Regex regex = A1SimpleRegex;

        StringBuilder sb = new();
        int lastIndex = 0;

        string shiftedWsName = shiftedRange.Worksheet.Name;
        foreach (Match match in regex.Matches(value).Cast<Match>())
        {
            string matchString = match.Value;
            int matchIndex = match.Index;
            if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
            {
                // Check that the match is not between quotes
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                string sheetName;
                bool useSheetName = false;
                if (matchString.Contains('!'))
                {
                    sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                    if (sheetName[0] == '\'')
                    {
                        sheetName = sheetName.Substring(1, sheetName.Length - 2);
                    }

                    useSheetName = true;
                }
                else
                {
                    sheetName = worksheetInAction.Name;
                }

                if (string.Compare(sheetName, shiftedWsName, true) == 0)
                {
                    string rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                    if (!A1ColumnRegex.IsMatch(rangeAddress))
                    {
                        IXLRange matchRange = worksheetInAction
                            .Workbook.Worksheet(sheetName)
                            .Range(rangeAddress);
                        if (
                            shiftedRange.RangeAddress.FirstAddress.RowNumber
                                <= matchRange.RangeAddress.LastAddress.RowNumber
                            && shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                                <= matchRange.RangeAddress.FirstAddress.ColumnNumber
                            && shiftedRange.RangeAddress.LastAddress.ColumnNumber
                                >= matchRange.RangeAddress.LastAddress.ColumnNumber
                        )
                        {
                            if (useSheetName)
                            {
                                sb.Append(sheetName.EscapeSheetName());
                                sb.Append('!');
                            }

                            if (A1RowRegex.IsMatch(rangeAddress))
                            {
                                string[] rows = rangeAddress.Split(':');
                                string row1String = rows[0];
                                string row2String = rows[1];
                                string row1;
                                if (row1String[0] == '$')
                                {
                                    row1 =
                                        "$"
                                        + (
                                            XlsxSharp.XLHelper.TrimRowNumber(
                                                int.Parse(row1String.Substring(1)) + rowsShifted
                                            )
                                        ).ToInvariantString();
                                }
                                else
                                {
                                    row1 = (
                                        XlsxSharp.XLHelper.TrimRowNumber(
                                            int.Parse(row1String) + rowsShifted
                                        )
                                    ).ToInvariantString();
                                }

                                string row2;
                                if (row2String[0] == '$')
                                {
                                    row2 =
                                        "$"
                                        + (
                                            XlsxSharp.XLHelper.TrimRowNumber(
                                                int.Parse(row2String.Substring(1)) + rowsShifted
                                            )
                                        ).ToInvariantString();
                                }
                                else
                                {
                                    row2 = (
                                        XlsxSharp.XLHelper.TrimRowNumber(
                                            int.Parse(row2String) + rowsShifted
                                        )
                                    ).ToInvariantString();
                                }

                                sb.Append(row1);
                                sb.Append(':');
                                sb.Append(row2);
                            }
                            else if (
                                shiftedRange.RangeAddress.FirstAddress.RowNumber
                                <= matchRange.RangeAddress.FirstAddress.RowNumber
                            )
                            {
                                if (rangeAddress.Contains(':'))
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XlsxSharp.XLHelper.TrimRowNumber(
                                                matchRange.RangeAddress.FirstAddress.RowNumber
                                                    + rowsShifted
                                            ),
                                            matchRange.RangeAddress.FirstAddress.ColumnLetter,
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn
                                        )
                                    );
                                    sb.Append(':');
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XlsxSharp.XLHelper.TrimRowNumber(
                                                matchRange.RangeAddress.LastAddress.RowNumber
                                                    + rowsShifted
                                            ),
                                            matchRange.RangeAddress.LastAddress.ColumnLetter,
                                            matchRange.RangeAddress.LastAddress.FixedRow,
                                            matchRange.RangeAddress.LastAddress.FixedColumn
                                        )
                                    );
                                }
                                else
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XlsxSharp.XLHelper.TrimRowNumber(
                                                matchRange.RangeAddress.FirstAddress.RowNumber
                                                    + rowsShifted
                                            ),
                                            matchRange.RangeAddress.FirstAddress.ColumnLetter,
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn
                                        )
                                    );
                                }
                            }
                            else
                            {
                                sb.Append(matchRange.RangeAddress.FirstAddress);
                                sb.Append(':');
                                sb.Append(
                                    new XLAddress(
                                        worksheetInAction,
                                        XlsxSharp.XLHelper.TrimRowNumber(
                                            matchRange.RangeAddress.LastAddress.RowNumber
                                                + rowsShifted
                                        ),
                                        matchRange.RangeAddress.LastAddress.ColumnLetter,
                                        matchRange.RangeAddress.LastAddress.FixedRow,
                                        matchRange.RangeAddress.LastAddress.FixedColumn
                                    )
                                );
                            }
                        }
                        else
                        {
                            sb.Append(matchString);
                        }
                    }
                    else
                    {
                        sb.Append(matchString);
                    }
                }
                else
                {
                    sb.Append(matchString);
                }
            }
            else
            {
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));
            }

            lastIndex = matchIndex + matchString.Length;
        }

        if (lastIndex < value.Length)
        {
            sb.Append(value.Substring(lastIndex));
        }

        return sb.ToString();
    }

    internal void ShiftFormulaColumns(XLRange shiftedRange, int columnsShifted) =>
        this.FormulaA1 = ShiftFormulaColumns(
            this.FormulaA1,
            this.Worksheet,
            shiftedRange,
            columnsShifted
        );

    internal static string ShiftFormulaColumns(
        string formulaA1,
        XLWorksheet worksheetInAction,
        XLRange shiftedRange,
        int columnsShifted
    )
    {
        if (string.IsNullOrWhiteSpace(formulaA1))
        {
            return string.Empty;
        }

        string value = formulaA1;

        Regex regex = A1SimpleRegex;

        StringBuilder sb = new();
        int lastIndex = 0;

        foreach (Match match in regex.Matches(value).Cast<Match>())
        {
            string matchString = match.Value;
            int matchIndex = match.Index;
            if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
            {
                // Check that the match is not between quotes
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                string sheetName;
                bool useSheetName = false;
                if (matchString.Contains('!'))
                {
                    sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                    if (sheetName[0] == '\'')
                    {
                        sheetName = sheetName.Substring(1, sheetName.Length - 2);
                    }

                    useSheetName = true;
                }
                else
                {
                    sheetName = worksheetInAction.Name;
                }

                if (string.Compare(sheetName, shiftedRange.Worksheet.Name, true) == 0)
                {
                    string rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                    if (!A1RowRegex.IsMatch(rangeAddress))
                    {
                        IXLRange matchRange = worksheetInAction
                            .Workbook.Worksheet(sheetName)
                            .Range(rangeAddress);

                        if (
                            shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                                <= matchRange.RangeAddress.LastAddress.ColumnNumber
                            && shiftedRange.RangeAddress.FirstAddress.RowNumber
                                <= matchRange.RangeAddress.FirstAddress.RowNumber
                            && shiftedRange.RangeAddress.LastAddress.RowNumber
                                >= matchRange.RangeAddress.LastAddress.RowNumber
                        )
                        {
                            if (useSheetName)
                            {
                                sb.Append(sheetName.EscapeSheetName());
                                sb.Append('!');
                            }

                            if (A1ColumnRegex.IsMatch(rangeAddress))
                            {
                                string[] columns = rangeAddress.Split(':');
                                string column1String = columns[0];
                                string column2String = columns[1];
                                string column1;
                                if (column1String[0] == '$')
                                {
                                    column1 =
                                        "$"
                                        + XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                                            XlsxSharp.XLHelper.GetColumnNumberFromLetter(
                                                column1String.Substring(1)
                                            ) + columnsShifted,
                                            true
                                        );
                                }
                                else
                                {
                                    column1 = XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                                        XlsxSharp.XLHelper.GetColumnNumberFromLetter(column1String)
                                            + columnsShifted,
                                        true
                                    );
                                }

                                string column2;
                                if (column2String[0] == '$')
                                {
                                    column2 =
                                        "$"
                                        + XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                                            XlsxSharp.XLHelper.GetColumnNumberFromLetter(
                                                column2String.Substring(1)
                                            ) + columnsShifted,
                                            true
                                        );
                                }
                                else
                                {
                                    column2 = XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                                        XlsxSharp.XLHelper.GetColumnNumberFromLetter(column2String)
                                            + columnsShifted,
                                        true
                                    );
                                }

                                sb.Append(column1);
                                sb.Append(':');
                                sb.Append(column2);
                            }
                            else if (
                                shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                                <= matchRange.RangeAddress.FirstAddress.ColumnNumber
                            )
                            {
                                if (rangeAddress.Contains(':'))
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.FirstAddress.RowNumber,
                                            XlsxSharp.XLHelper.TrimColumnNumber(
                                                matchRange.RangeAddress.FirstAddress.ColumnNumber
                                                    + columnsShifted
                                            ),
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn
                                        )
                                    );
                                    sb.Append(':');
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.LastAddress.RowNumber,
                                            XlsxSharp.XLHelper.TrimColumnNumber(
                                                matchRange.RangeAddress.LastAddress.ColumnNumber
                                                    + columnsShifted
                                            ),
                                            matchRange.RangeAddress.LastAddress.FixedRow,
                                            matchRange.RangeAddress.LastAddress.FixedColumn
                                        )
                                    );
                                }
                                else
                                {
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.FirstAddress.RowNumber,
                                            XlsxSharp.XLHelper.TrimColumnNumber(
                                                matchRange.RangeAddress.FirstAddress.ColumnNumber
                                                    + columnsShifted
                                            ),
                                            matchRange.RangeAddress.FirstAddress.FixedRow,
                                            matchRange.RangeAddress.FirstAddress.FixedColumn
                                        )
                                    );
                                }
                            }
                            else
                            {
                                sb.Append(matchRange.RangeAddress.FirstAddress);
                                sb.Append(':');
                                sb.Append(
                                    new XLAddress(
                                        worksheetInAction,
                                        matchRange.RangeAddress.LastAddress.RowNumber,
                                        XlsxSharp.XLHelper.TrimColumnNumber(
                                            matchRange.RangeAddress.LastAddress.ColumnNumber
                                                + columnsShifted
                                        ),
                                        matchRange.RangeAddress.LastAddress.FixedRow,
                                        matchRange.RangeAddress.LastAddress.FixedColumn
                                    )
                                );
                            }
                        }
                        else
                        {
                            sb.Append(matchString);
                        }
                    }
                    else
                    {
                        sb.Append(matchString);
                    }
                }
                else
                {
                    sb.Append(matchString);
                }
            }
            else
            {
                sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));
            }

            lastIndex = matchIndex + matchString.Length;
        }

        if (lastIndex < value.Length)
        {
            sb.Append(value.Substring(lastIndex));
        }

        return sb.ToString();
    }

    private XLCell CellShift(int rowsToShift, int columnsToShift) =>
        this.Worksheet.Cell(this._rowNumber + rowsToShift, this._columnNumber + columnsToShift);

    #region XLCell Above

    IXLCell IXLCell.CellAbove() => this.CellAbove();

    IXLCell IXLCell.CellAbove(int step) => this.CellAbove(step);

    public XLCell CellAbove() => this.CellAbove(1);

    public XLCell CellAbove(int step) => this.CellShift(step * -1, 0);

    #endregion XLCell Above

    #region XLCell Below

    IXLCell IXLCell.CellBelow() => this.CellBelow();

    IXLCell IXLCell.CellBelow(int step) => this.CellBelow(step);

    public XLCell CellBelow() => this.CellBelow(1);

    public XLCell CellBelow(int step) => this.CellShift(step, 0);

    #endregion XLCell Below

    #region XLCell Left

    IXLCell IXLCell.CellLeft() => this.CellLeft();

    IXLCell IXLCell.CellLeft(int step) => this.CellLeft(step);

    public XLCell CellLeft() => this.CellLeft(1);

    public XLCell CellLeft(int step) => this.CellShift(0, step * -1);

    #endregion XLCell Left

    #region XLCell Right

    IXLCell IXLCell.CellRight() => this.CellRight();

    IXLCell IXLCell.CellRight(int step) => this.CellRight(step);

    public XLCell CellRight() => this.CellRight(1);

    public XLCell CellRight(int step) => this.CellShift(0, step);

    #endregion XLCell Right

    public bool HasFormula => this.Formula is not null;

    public bool HasArrayFormula => this.Formula?.Type == FormulaType.Array;

    public IXLRangeAddress FormulaReference
    {
        get
        {
            if (this.Formula is null)
            {
                return null;
            }

            Area range = this.Formula.Range;
            if (range == default)
            {
                return null;
            }

            return XLRangeAddress.FromSheetRange(this.Worksheet, range);
        }
        set
        {
            if (this.Formula is null)
            {
                throw new ArgumentException("Cell doesn't contain a formula.");
            }

            if (value is null)
            {
                this.Formula.Range = default;
                return;
            }

            if (value.Worksheet is not null && this.Worksheet != value.Worksheet)
            {
                throw new ArgumentException(
                    "The reference worksheet must be same as worksheet of the cell or null."
                );
            }

            this.Formula.Range = Area.FromRangeAddress(value);
        }
    }

    public IXLRange CurrentRegion => this.Worksheet.Range(this.FindCurrentRegion());

    private IXLRangeAddress FindCurrentRegion()
    {
        XLWorksheet sheet = this.Worksheet;

        int minRow = this._rowNumber;
        int minCol = this._columnNumber;
        int maxRow = this._rowNumber;
        int maxCol = this._columnNumber;

        bool hasRegionExpanded;

        do
        {
            hasRegionExpanded = false;

            int borderMinRow = Math.Max(minRow - 1, XlsxSharp.XLHelper.MinRowNumber);
            int borderMaxRow = Math.Min(maxRow + 1, XlsxSharp.XLHelper.MaxRowNumber);
            int borderMinColumn = Math.Max(minCol - 1, XlsxSharp.XLHelper.MinColumnNumber);
            int borderMaxColumn = Math.Min(maxCol + 1, XlsxSharp.XLHelper.MaxColumnNumber);

            if (
                minCol > XlsxSharp.XLHelper.MinColumnNumber
                && !IsVerticalBorderBlank(sheet, borderMinColumn, borderMinRow, borderMaxRow)
            )
            {
                hasRegionExpanded = true;
                minCol = borderMinColumn;
            }

            if (
                maxCol < XlsxSharp.XLHelper.MaxColumnNumber
                && !IsVerticalBorderBlank(sheet, borderMaxColumn, borderMinRow, borderMaxRow)
            )
            {
                hasRegionExpanded = true;
                maxCol = borderMaxColumn;
            }

            if (
                minRow > XlsxSharp.XLHelper.MinRowNumber
                && !IsHorizontalBorderBlank(sheet, borderMinRow, borderMinColumn, borderMaxColumn)
            )
            {
                hasRegionExpanded = true;
                minRow = borderMinRow;
            }

            if (
                maxRow < XlsxSharp.XLHelper.MaxRowNumber
                && !IsHorizontalBorderBlank(sheet, borderMaxRow, borderMinColumn, borderMaxColumn)
            )
            {
                hasRegionExpanded = true;
                maxRow = borderMaxRow;
            }
        } while (hasRegionExpanded);

        return new XLRangeAddress(
            new XLAddress(sheet, minRow, minCol, false, false),
            new XLAddress(sheet, maxRow, maxCol, false, false)
        );

        static bool IsVerticalBorderBlank(
            XLWorksheet sheet,
            int borderColumn,
            int borderMinRow,
            int borderMaxRow
        )
        {
            for (int row = borderMinRow; row <= borderMaxRow; row++)
            {
                XLCell verticalBorderCell = sheet.Cell(row, borderColumn);
                if (!verticalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
                {
                    return false;
                }
            }

            return true;
        }

        static bool IsHorizontalBorderBlank(
            XLWorksheet sheet,
            int borderRow,
            int borderMinColumn,
            int borderMaxColumn
        )
        {
            for (int col = borderMinColumn; col <= borderMaxColumn; col++)
            {
                XLCell horizontalBorderCell = sheet.Cell(borderRow, col);
                if (!horizontalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
                {
                    return false;
                }
            }

            return true;
        }
    }

    internal bool IsInferiorMergedCell() =>
        this.IsMerged() && !this.Address.Equals(this.MergedRange().RangeAddress.FirstAddress);

    /// <summary>
    /// Get glyph bounding boxes for each grapheme in the text. Box size is determined according to
    /// the font of a grapheme. New lines are represented as default (all dimensions zero) box.
    /// A line without any text (i.e. contains only new line) should be represented by a box
    /// with zero advance width, but with a line height of corresponding font.
    /// </summary>
    /// <param name="engine">Engine used to determine box size.</param>
    /// <param name="dpi">DPI used to determine size of glyphs.</param>
    /// <param name="output">List where items are added.</param>
    internal void GetGlyphBoxes(IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
    {
        XLImmutableRichText richText = this.SliceRichText;
        if (richText is not null)
        {
            foreach (XLImmutableRichText.RichTextRun richTextRun in richText.Runs)
            {
                string text = richText.GetRunText(richTextRun);
                IXLFontBase font = richTextRun.Font.ToFontBase();
                AddGlyphs(text, font, engine, dpi, output);
            }
        }
        else
        {
            string text = this.GetFormattedString();
            AddGlyphs(text, this.Style.Font, engine, dpi, output);
        }

        static void AddGlyphs(
            string text,
            IXLFontBase font,
            IXLGraphicEngine engine,
            Dpi dpi,
            List<GlyphBox> output
        )
        {
            Span<int> zeroWidthJoiner = [0x200D];
            bool prevWasNewLine = false;
            int[] graphemeStarts = StringInfo.ParseCombiningCharacters(text);
            ReadOnlySpan<char> textSpan = text.AsSpan();

            // If we have more than 1 code unit per grapheme, the code units can
            // be distributed through multiple grapheme. In the worst case, all extra
            // code units are in exactly one grapheme -> allocate buffer of that size.
            Span<int> codePointsBuffer = stackalloc int[1 + text.Length - graphemeStarts.Length];
            for (int i = 0; i < graphemeStarts.Length; ++i)
            {
                int startIdx = graphemeStarts[i];
                ReadOnlySpan<char> slice = textSpan.Slice(startIdx);
                if (slice.TrySliceNewLine(out int eolLen))
                {
                    i += eolLen - 1;
                    if (prevWasNewLine)
                    {
                        // If there are consecutive new lines, we need height of new the lines between them
                        GlyphBox box = engine.GetGlyphBox(zeroWidthJoiner, font, dpi);
                        output.Add(box);
                    }

                    output.Add(GlyphBox.LineBreak);
                    prevWasNewLine = true;
                }
                else
                {
                    ReadOnlySpan<char> codeUnits =
                        i + 1 < graphemeStarts.Length
                            ? textSpan.Slice(startIdx, graphemeStarts[i + 1] - startIdx)
                            : textSpan.Slice(startIdx);
                    int count = codeUnits.ToCodePoints(codePointsBuffer);
                    ReadOnlySpan<int> grapheme = codePointsBuffer.Slice(0, count);
                    GlyphBox box = engine.GetGlyphBox(grapheme, font, dpi);
                    output.Add(box);
                    prevWasNewLine = false;
                }
            }
        }
    }

    public override int GetHashCode() => HashCode.Combine(this.Point, this.Worksheet);

    public override bool Equals(object obj) =>
        obj is XLCell cell && cell.Worksheet == this.Worksheet && cell.Point == this.Point;
}
