using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Excel;

/// <summary>
/// API object to modify properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// The methods and properties create a modified formats and the formats are registered
/// in the <see cref="XLWorkbookStyles"/>.
/// </summary>
internal partial class XLCellFormat
{
    private readonly XLWorkbook _workbook;
    private readonly Hierarchy _formatValue;

    private XLCellFormat(XLWorkbook workbook, Hierarchy formatValue)
    {
        this._workbook = workbook;
        this._formatValue = formatValue;
    }

    internal XLNumberCellFormat NumberFormat => new(this);

    internal XLFontCellFormat Font => new(this);

    internal XLFillCellFormat Fill => new(this);

    internal XLBorderCellFormat Border => new(this);

    internal XLAlignmentCellFormat Alignment => new(this);

    internal XLProtectionCellFormat Protection => new(this);

    internal bool IncludeQuotePrefix
    {
        get => this.Resolve(static format => format.IncludeQuotePrefix);
        set =>
            this.ModifyFormat(
                (format, includeQuotePrefix) =>
                    format with
                    {
                        IncludeQuotePrefix = includeQuotePrefix,
                    },
                value
            );
    }

    /// <summary>
    /// Cell areas in a workbook that should be updated when format is changed, e.g. when we have
    /// a format API object for a row container, the area are all cells of the row. It must be
    /// an area, so we can satisfy the <see cref="IXLBorder.OutsideBorder"/> and
    /// <see cref="IXLBorder.InsideBorder"/> property setters.
    /// </summary>
    private IReadOnlyList<SheetArea> Areas { get; init; } = Array.Empty<SheetArea>();

    /// <summary>
    /// Formatting is updated for used cells within these areas. Unused cells are ignored.
    /// </summary>
    private IReadOnlyList<SheetArea> UsedAreas { get; init; } = Array.Empty<SheetArea>();

    /// <summary>
    /// Formatting is updated for these columns. This doesn't update cells within the columns, only
    /// the columns themselves. The values are unique columns, sorted by column number in ascending
    /// order.
    /// </summary>
    private XLColumnArea[] Columns { get; init; } = [];

    /// <summary>
    /// Formatting is updated for these rows. This doesn't update cells within the rows, only
    /// the rows themselves. The values are unique rows, sorted by row number in ascending order.
    /// </summary>
    private XLRowArea[] Rows { get; init; } = [];

    /// <summary>
    /// Formatting is updated for these worksheets. This doesn't update cells within the sheets, only
    /// the sheets and materialized rows and columns of the sheets.
    /// </summary>
    private IReadOnlyList<string> Worksheets { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Should the formatting be updated for the default format of a workbook (plus cascade to all
    /// formats below, containers and areas).
    /// </summary>
    private bool DefaultFormat { get; init; }

    /// <summary>
    /// A flag indicating API object is for XLCells. Unlike other range API objects, the XLCells
    /// has a non-standard outside/inside borders behavior.
    /// </summary>
    private bool IsCells { get; init; }

    public override bool Equals(object? obj)
    {
        return obj is IXLStyle other && (this as IEquatable<IXLStyle>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal static XLCellFormat ForCell(XLCell cell)
    {
        XLWorkbook workbook = cell.Worksheet.Workbook;
        string sheetName = cell.Worksheet.Name;
        Point cellPoint = cell.Point;
        Hierarchy formatValue = new(
            workbook,
            sheetName,
            cellPoint.Column,
            cellPoint.Row,
            cellPoint
        );
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = new[] { new SheetArea(sheetName, new Area(cellPoint)) },
        };
    }

    internal static XLCellFormat ForColumn(XLColumn column)
    {
        XLWorkbook workbook = column.Worksheet.Workbook;
        XLColumnArea columnArea = column.Area;
        Hierarchy formatValue = new(workbook, columnArea.Name, columnArea.ColumNumber, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { columnArea.Area },
            Columns = [columnArea],
        };
    }

    internal static XLCellFormat ForColumns(
        XLWorkbook workbook,
        XLWorksheet? formatValueSheet,
        IEnumerable<XLColumn> columns
    )
    {
        XLColumnArea[] columnAreas =
        [
            .. columns.Select(x => x.Area).Distinct().OrderBy(x => x.ColumNumber),
        ];
        Hierarchy formatValue = new(workbook, formatValueSheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Columns = columnAreas,
            UsedAreas = columnAreas.Select(x => x.Area).ToArray(),
        };
    }

    internal static XLCellFormat ForRow(XLRow row)
    {
        XLWorkbook workbook = row.Worksheet.Workbook;
        XLRowArea rowArea = row.Area;
        Hierarchy formatValue = new(workbook, rowArea.Name, null, rowArea.RowNumber, null);
        return new XLCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { rowArea.Area },
            Rows = [rowArea],
        };
    }

    internal static XLCellFormat ForRows(
        XLWorkbook workbook,
        XLWorksheet? formatValueSheet,
        IEnumerable<XLRow> rows
    )
    {
        XLRowArea[] rowAreas = [.. rows.Select(x => x.Area).Distinct().OrderBy(x => x.RowNumber)];
        Hierarchy formatValue = new(workbook, formatValueSheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Rows = rowAreas,
            UsedAreas = rowAreas.Select(x => x.Area).ToArray(),
        };
    }

    internal static XLCellFormat ForWorksheet(XLWorksheet worksheet)
    {
        XLWorkbook workbook = worksheet.Workbook;
        Hierarchy formatValue = new(workbook, worksheet.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { worksheet.Area },
            Worksheets = new[] { worksheet.Name },
        };
    }

    internal static XLCellFormat ForWorkbook(XLWorkbook workbook)
    {
        Hierarchy formatValue = new(workbook, null, null, null, null);
        return new XLCellFormat(workbook, formatValue) { DefaultFormat = true };
    }

    internal static XLCellFormat ForAreas(
        XLWorkbook workbook,
        IReadOnlyList<SheetArea> areas,
        XLWorksheet? sheet
    )
    {
        Hierarchy formatValue = new(workbook, sheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue) { Areas = areas };
    }

    internal static XLCellFormat ForCells(
        XLWorkbook workbook,
        IReadOnlyList<SheetArea> areas,
        XLWorksheet? sheet
    )
    {
        Hierarchy formatValue = new(workbook, sheet?.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue) { Areas = areas, IsCells = true };
    }

    internal static XLCellFormat ForRange(XLWorksheet sheet, XLRangeAddress rangeAddress)
    {
        XLWorkbook workbook = sheet.Workbook;
        Hierarchy formatValue = new(workbook, sheet.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue)
        {
            Areas = new[] { SheetArea.From(rangeAddress) },
        };
    }

    internal static XLCellFormat ForTableRows(XLWorksheet sheet, SheetArea[] rowAreas)
    {
        XLWorkbook workbook = sheet.Workbook;
        Hierarchy formatValue = new(workbook, sheet.Name, null, null, null);
        return new XLCellFormat(workbook, formatValue) { Areas = rowAreas };
    }

    internal T Resolve<T>(Func<XLCellFormatValue, T> selector)
    {
        XLCellFormatValue format = this._formatValue.Resolve();
        return selector(format);
    }

    internal void ModifyFormat<TProperty>(
        Func<XLCellFormatValue, TProperty, XLCellFormatValue> modifyFormat,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(format =>
            styles.GetRegisteredCellFormat(format, cellFormat => modifyFormat(cellFormat, value))
        );
    }

    // TODO Styles: Move modification methods of each component to the XLCellCollection. Modification
    // of component should always update CustomFormat and to make sure that is done, it should be done
    // in a one place.
    internal void ModifyNumberFormat(XLNumberFormat numberFormat)
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(format =>
        {
            XLNumberFormat modifiedNumberFormat = styles.RegisterNumberFormat(numberFormat);
            XLCellFormatValue modifiedFormat = styles.GetRegisteredCellFormat(
                format,
                cellFormat =>
                    cellFormat with
                    {
                        NumberFormat = modifiedNumberFormat,
                        CustomFormat = format.CustomFormat | CellFormatComponents.NumberFormat,
                    }
            );
            return modifiedFormat;
        });
    }

    internal void ModifyFont<TProperty>(
        Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(format =>
        {
            XLFontFormatValue modifiedFont = styles.GetRegisteredFontFormat(
                format.Font,
                font => modifyFont(font, value)
            );
            XLCellFormatValue modifiedFormat = styles.GetRegisteredCellFormat(
                format,
                cellFormat =>
                    cellFormat with
                    {
                        Font = modifiedFont,
                        CustomFormat = format.CustomFormat | CellFormatComponents.Font,
                    }
            );
            return modifiedFormat;
        });
    }

    internal void ModifyFill<TProperty>(
        Func<XLFillFormatValue, TProperty, XLFillFormatValue> modifyFill,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(format =>
        {
            XLFillFormatValue modifiedFill = styles.GetRegisteredFillFormat(
                format.Fill,
                fill => modifyFill(fill, value)
            );
            XLCellFormatValue modifiedFormat = styles.GetRegisteredCellFormat(
                format,
                cellFormat =>
                    cellFormat with
                    {
                        Fill = modifiedFill,
                        CustomFormat = format.CustomFormat | CellFormatComponents.Fill,
                    }
            );
            return modifiedFormat;
        });
    }

    internal void ModifyBorder<TProperty>(
        Func<XLBorderFormatValue, TProperty, XLBorderFormatValue> modifyBorder,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(GetModifyBorderFunc(border => modifyBorder(border, value), styles));
    }

    internal void ModifyAlignment<TProperty>(
        Func<XLAlignmentFormatValue, TProperty, XLAlignmentFormatValue> modifyAlignment,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(format =>
        {
            XLAlignmentFormatValue modifiedAlignment = styles.RegisterAlignmentFormat(
                modifyAlignment(format.Alignment, value)
            );
            XLCellFormatValue modifiedFormat = styles.GetRegisteredCellFormat(
                format,
                cellFormat =>
                    cellFormat with
                    {
                        Alignment = modifiedAlignment,
                        CustomFormat = format.CustomFormat | CellFormatComponents.Alignment,
                    }
            );
            return modifiedFormat;
        });
    }

    internal void ModifyProtection<TProperty>(
        Func<XLProtectionFormatValue, TProperty, XLProtectionFormatValue> modifyProtection,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        this.Modify(format =>
        {
            XLProtectionFormatValue modifiedProtection = styles.RegisterProtectionFormat(
                modifyProtection(format.Protection, value)
            );
            XLCellFormatValue modifiedFormat = styles.GetRegisteredCellFormat(
                format,
                cellFormat =>
                    cellFormat with
                    {
                        Protection = modifiedProtection,
                        CustomFormat = format.CustomFormat | CellFormatComponents.Protection,
                    }
            );
            return modifiedFormat;
        });
    }

    internal void ModifyOuterBorder<TProperty>(
        Func<XLBorderLine, TProperty, XLBorderLine> modify,
        TProperty value
    )
    {
        XLWorkbookStyles styles = this._workbook.Styles;
        if (this.IsCells)
        {
            Func<XLCellFormatValue, XLCellFormatValue> setAll = GetModifyBorderFunc(
                border =>
                    border with
                    {
                        Left = modify(border.Left, value),
                        Top = modify(border.Top, value),
                        Right = modify(border.Right, value),
                        Bottom = modify(border.Bottom, value),
                    },
                styles
            );
            foreach ((string sheetName, Area area) in this.Areas)
            {
                if (!this._workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
                {
                    continue;
                }

                ApplyToAll(area, setAll, worksheet);
            }

            return;
        }

        // Change only top and bottom border of a row. The style is used by non-materialized cells
        // in a row and will be used by non-materialized cells in a row. Same applies to columns.
        Func<XLCellFormatValue, XLCellFormatValue> setTopAndBottom = GetModifyBorderFunc(
            border =>
                border with
                {
                    Top = modify(border.Top, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );
        this.ModifyRowsBorder(this.Rows, setTopAndBottom);

        Func<XLCellFormatValue, XLCellFormatValue> setLeftAndRight = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Right = modify(border.Right, value),
                },
            styles
        );
        this.ModifyColumnsBorder(this.Columns, setLeftAndRight);

        // A normal path for range API object (except XLCells). Set outer border to areas.
        // Don't use UsedAreas, they are for columns/rows. Worksheet doesn't have an outer border.
        Func<XLCellFormatValue, XLCellFormatValue> setLeft = GetModifyBorderFunc(
            border => border with { Left = modify(border.Left, value) },
            styles
        );
        Func<XLCellFormatValue, XLCellFormatValue> setTop = GetModifyBorderFunc(
            border => border with { Top = modify(border.Top, value) },
            styles
        );
        Func<XLCellFormatValue, XLCellFormatValue> setRight = GetModifyBorderFunc(
            border => border with { Right = modify(border.Right, value) },
            styles
        );
        Func<XLCellFormatValue, XLCellFormatValue> setBottom = GetModifyBorderFunc(
            border => border with { Bottom = modify(border.Bottom, value) },
            styles
        );
        foreach (SheetArea area in this.Areas)
        {
            if (!this._workbook.TryGetWorksheet(area.Name, out XLWorksheet worksheet))
            {
                continue;
            }

            FormatResolver formatResolver = new(worksheet);
            XLCellsCollection cellsCollection = worksheet.Internals.CellsCollection;

            // Left side
            Area left = area.Area.SliceFromLeft(1);
            cellsCollection.ApplyFormatOnAll(left, setLeft, formatResolver.Resolve);

            // Top side
            Area top = area.Area.SliceFromTop(1);
            cellsCollection.ApplyFormatOnAll(top, setTop, formatResolver.Resolve);

            // Right side
            Area right = area.Area.SliceFromRight(1);
            cellsCollection.ApplyFormatOnAll(right, setRight, formatResolver.Resolve);

            // Bottom side
            Area bottom = area.Area.SliceFromBottom(1);
            cellsCollection.ApplyFormatOnAll(bottom, setBottom, formatResolver.Resolve);
        }
    }

    internal void ModifyInnerBorder<TProperty>(
        Func<XLBorderLine, TProperty, XLBorderLine> modify,
        TProperty value
    )
    {
        // Shortcut for XLCells - it has no inner borders
        if (this.IsCells)
        {
            return;
        }

        XLWorkbookStyles styles = this._workbook.Styles;
        this.ModifyInsideBordersOfRows(styles, modify, value);
        this.ModifyInsideBordersOfColumns(styles, modify, value);

        Func<XLCellFormatValue, XLCellFormatValue> setLeft = GetModifyBorderFunc(
            border => border with { Left = modify(border.Left, value) },
            styles
        );
        Func<XLCellFormatValue, XLCellFormatValue> setTop = GetModifyBorderFunc(
            border => border with { Top = modify(border.Top, value) },
            styles
        );
        Func<XLCellFormatValue, XLCellFormatValue> setRight = GetModifyBorderFunc(
            border => border with { Right = modify(border.Right, value) },
            styles
        );
        Func<XLCellFormatValue, XLCellFormatValue> setBottom = GetModifyBorderFunc(
            border => border with { Bottom = modify(border.Bottom, value) },
            styles
        );
        foreach ((string sheetName, Area area) in this.Areas)
        {
            if (!this._workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
            {
                continue;
            }

            FormatResolver formatResolver = new(worksheet);
            XLCellsCollection cellsCollection = worksheet.Internals.CellsCollection;

            // Setting line from both sides is not super useful, but keeps internal state consistent.
            if (area.Width > 1)
            {
                cellsCollection.ApplyFormatOnAll(
                    area.SliceFromLeft(area.Width - 1),
                    setRight,
                    formatResolver.Resolve
                );
                cellsCollection.ApplyFormatOnAll(
                    area.SliceFromRight(area.Width - 1),
                    setLeft,
                    formatResolver.Resolve
                );
            }

            if (area.Height > 1)
            {
                cellsCollection.ApplyFormatOnAll(
                    area.SliceFromTop(area.Height - 1),
                    setBottom,
                    formatResolver.Resolve
                );
                cellsCollection.ApplyFormatOnAll(
                    area.SliceFromBottom(area.Height - 1),
                    setTop,
                    formatResolver.Resolve
                );
            }
        }
    }

    private static Func<XLCellFormatValue, XLCellFormatValue> GetModifyBorderFunc(
        Func<XLBorderFormatValue, XLBorderFormatValue> modifyBorder,
        XLWorkbookStyles styles
    )
    {
        return format =>
        {
            XLBorderFormatValue modifiedBorder = styles.GetRegisteredBorderFormat(
                format.Border,
                border =>
                {
                    XLBorderFormatValue modified = modifyBorder(border);

                    // Per original behavior, the non-visible border can't hold color state, e.g. when
                    // a border is set to from Thin to None and later changed back to Thick, it
                    // shouldn't remember the original color.
                    // That is not how Excel behaves and it makes everything harder (e.g. user can't
                    // set the border color first and then border style), but it is what it is.
                    if (!modified.Left.IsVisible)
                    {
                        modified = modified with { Left = XLBorderLine.None };
                    }

                    if (!modified.Top.IsVisible)
                    {
                        modified = modified with { Top = XLBorderLine.None };
                    }

                    if (!modified.Right.IsVisible)
                    {
                        modified = modified with { Right = XLBorderLine.None };
                    }

                    if (!modified.Bottom.IsVisible)
                    {
                        modified = modified with { Bottom = XLBorderLine.None };
                    }

                    if (!modified.Diagonal.IsVisible)
                    {
                        modified = modified with { Diagonal = XLBorderLine.None };
                    }

                    return modified;
                }
            );
            XLCellFormatValue modifiedFormat = styles.GetRegisteredCellFormat(
                format,
                cellFormat =>
                    cellFormat with
                    {
                        Border = modifiedBorder,
                        CustomFormat = format.CustomFormat | CellFormatComponents.Border,
                    }
            );
            return modifiedFormat;
        };
    }

    private void ModifyInsideBordersOfRows<TProperty>(
        XLWorkbookStyles styles,
        Func<XLBorderLine, TProperty, XLBorderLine> modify,
        TProperty value
    )
    {
        // For a single row, only the left are right border are counted as "inside". The top and bottom border touch the outside.
        Func<XLCellFormatValue, XLCellFormatValue> setLeftAndRight = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Right = modify(border.Right, value),
                },
            styles
        );

        // For multi-row rowspan, there are three different patterns:
        // Multi-row rowspan - top row
        Func<XLCellFormatValue, XLCellFormatValue> setLeftRightBottom = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Right = modify(border.Right, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );

        // Multi-row rowspan - center rows. There isn't a center row in 2-row rowspan
        Func<XLCellFormatValue, XLCellFormatValue> setAll = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Top = modify(border.Top, value),
                    Right = modify(border.Right, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );

        // Multi-row rowspan - bottom row
        Func<XLCellFormatValue, XLCellFormatValue> setLeftTopRight = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Top = modify(border.Top, value),
                    Right = modify(border.Right, value),
                },
            styles
        );

        // Set border for each rowspan
        for (int i = 0; i < this.Rows.Length; ++i)
        {
            // Find rowspan as a sequence of consecutive rows
            int startIndex = i;
            int endIndex = i;
            while (
                endIndex + 1 < this.Rows.Length
                && this.Rows[endIndex + 1].RowNumber == this.Rows[endIndex].RowNumber + 1
            )
            {
                endIndex += 1;
            }

            i = endIndex;
            int rowSpanHeight = endIndex - startIndex + 1;
            if (rowSpanHeight > 1)
            {
                this.ModifyRowsBorder(this.Rows.AsSpan(startIndex, 1), setLeftRightBottom);
                this.ModifyRowsBorder(this.Rows.AsSpan(startIndex + 1, rowSpanHeight - 2), setAll);
                this.ModifyRowsBorder(this.Rows.AsSpan(endIndex, 1), setLeftTopRight);
            }
            else
            {
                this.ModifyRowsBorder(this.Rows.AsSpan(startIndex, 1), setLeftAndRight);
            }
        }
    }

    private void ModifyRowsBorder(
        ReadOnlySpan<XLRowArea> rows,
        Func<XLCellFormatValue, XLCellFormatValue> modifyBorder
    )
    {
        foreach (XLRowArea rowArea in rows)
        {
            if (!this._workbook.TryGetWorksheet(rowArea.Name, out XLWorksheet worksheet))
            {
                continue;
            }

            // Row style is used by non-materialized cells in a row...
            XLRow row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, modifyBorder, worksheet);

            // ... and materialized cells in a row have format explicitly set.
            FormatResolver formatResolver = new(worksheet);
            XLCellsCollection cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(
                rowArea.Area.Area,
                modifyBorder,
                formatResolver.Resolve
            );
        }
    }

    private void ModifyInsideBordersOfColumns<TProperty>(
        XLWorkbookStyles styles,
        Func<XLBorderLine, TProperty, XLBorderLine> modify,
        TProperty value
    )
    {
        // For a single column, only the top are bottom border are counted as "inside". The left and right border touch the outside.
        Func<XLCellFormatValue, XLCellFormatValue> setTopAndBottom = GetModifyBorderFunc(
            border =>
                border with
                {
                    Top = modify(border.Top, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );

        // For multi-column colspan, there are three different patterns:
        // Multi-column colspan - left column
        Func<XLCellFormatValue, XLCellFormatValue> setTopRightBottom = GetModifyBorderFunc(
            border =>
                border with
                {
                    Top = modify(border.Top, value),
                    Right = modify(border.Right, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );

        // Multi-column colspan - center columns. There isn't a center column in 2-column colspan
        Func<XLCellFormatValue, XLCellFormatValue> setAll = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Top = modify(border.Top, value),
                    Right = modify(border.Right, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );

        // Multi-column colspan - right column
        Func<XLCellFormatValue, XLCellFormatValue> setLeftTopBottom = GetModifyBorderFunc(
            border =>
                border with
                {
                    Left = modify(border.Left, value),
                    Top = modify(border.Top, value),
                    Bottom = modify(border.Bottom, value),
                },
            styles
        );

        // Set border for each colspan
        for (int i = 0; i < this.Columns.Length; ++i)
        {
            // Find colspan as a sequence of consecutive columns
            int startIndex = i;
            int endIndex = i;
            while (
                endIndex + 1 < this.Columns.Length
                && this.Columns[endIndex + 1].ColumNumber == this.Columns[endIndex].ColumNumber + 1
            )
            {
                endIndex += 1;
            }

            i = endIndex;
            int colspanWidth = endIndex - startIndex + 1;
            if (colspanWidth > 1)
            {
                this.ModifyColumnsBorder(this.Columns.AsSpan(startIndex, 1), setTopRightBottom);
                this.ModifyColumnsBorder(
                    this.Columns.AsSpan(startIndex + 1, colspanWidth - 2),
                    setAll
                );
                this.ModifyColumnsBorder(this.Columns.AsSpan(endIndex, 1), setLeftTopBottom);
            }
            else
            {
                this.ModifyColumnsBorder(this.Columns.AsSpan(startIndex, 1), setTopAndBottom);
            }
        }
    }

    private void ModifyColumnsBorder(
        ReadOnlySpan<XLColumnArea> columns,
        Func<XLCellFormatValue, XLCellFormatValue> modifyBorder
    )
    {
        foreach (XLColumnArea columnArea in columns)
        {
            if (!this._workbook.TryGetWorksheet(columnArea.Name, out XLWorksheet worksheet))
            {
                continue;
            }

            // Column style is used by non-materialized cells in a column...
            XLColumn column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, modifyBorder, worksheet);

            // ... and materialized cells in a column have format explicitly set.
            FormatResolver formatResolver = new(worksheet);
            XLCellsCollection cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(
                columnArea.Area.Area,
                modifyBorder,
                formatResolver.Resolve
            );
        }
    }

    private void Modify(Func<XLCellFormatValue, XLCellFormatValue> modifyFormat)
    {
        // TODO Styles: Deal with cross points
        XLWorkbookStyles styles = this._workbook.Styles;
        if (this.DefaultFormat)
        {
            styles.DefaultFormat = modifyFormat(styles.DefaultFormat);
            foreach (XLWorksheet worksheet in this._workbook.WorksheetsInternal)
            {
                ApplyToWorksheet(worksheet, modifyFormat, styles);
            }
        }

        foreach (string sheetName in this.Worksheets)
        {
            if (!this._workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
            {
                continue;
            }

            ApplyToWorksheet(worksheet, modifyFormat, styles);
        }

        foreach (XLColumnArea columnArea in this.Columns)
        {
            if (!this._workbook.TryGetWorksheet(columnArea.Name, out XLWorksheet worksheet))
            {
                continue;
            }

            XLColumn column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, modifyFormat, worksheet);
        }

        foreach (XLRowArea rowArea in this.Rows)
        {
            if (!this._workbook.TryGetWorksheet(rowArea.Name, out XLWorksheet worksheet))
            {
                continue;
            }

            XLRow row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, modifyFormat, worksheet);
        }

        foreach ((string sheetName, Area area) in this.UsedAreas)
        {
            if (!this._workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
            {
                continue;
            }

            ApplyToUsed(area, modifyFormat, worksheet);
        }

        foreach ((string sheetName, Area area) in this.Areas)
        {
            if (!this._workbook.TryGetWorksheet(sheetName, out XLWorksheet worksheet))
            {
                continue;
            }

            ApplyToAll(area, modifyFormat, worksheet);
        }
    }

    private static void ApplyToWorksheet(
        XLWorksheet worksheet,
        Func<XLCellFormatValue, XLCellFormatValue> modifyFormat,
        XLWorkbookStyles styles
    )
    {
        XLCellFormatValue originalFormat = worksheet.FormatValue ?? styles.DefaultFormat;
        XLCellFormatValue modifiedFormat = modifyFormat(originalFormat);
        worksheet.FormatValue = modifiedFormat;

        ICollection<XLColumn> columns = worksheet.Internals.ColumnsCollection.Values;
        foreach (XLColumn column in columns)
        {
            ApplyColRowFormat(column, modifyFormat, worksheet);
        }

        ICollection<XLRow> rows = worksheet.Internals.RowsCollection.Values;
        foreach (XLRow row in rows)
        {
            ApplyColRowFormat(row, modifyFormat, worksheet);
        }

        ApplyToUsed(Area.Full, modifyFormat, worksheet);
    }

    private static void ApplyColRowFormat(
        IXLFormatContainer rowOrCol,
        Func<XLCellFormatValue, XLCellFormatValue> modifyFormat,
        XLWorksheet worksheet
    )
    {
        if (rowOrCol.FormatValue is not { } originalFormat)
        {
            originalFormat = worksheet.FormatValue ?? worksheet.Workbook.Styles.DefaultFormat;
        }

        rowOrCol.FormatValue = modifyFormat(originalFormat);
    }

    private static void ApplyToUsed(
        Area area,
        Func<XLCellFormatValue, XLCellFormatValue> modifyFormat,
        XLWorksheet worksheet
    )
    {
        FormatResolver formatResolver = new(worksheet);
        XLCellsCollection cellsCollection = worksheet.Internals.CellsCollection;
        cellsCollection.ApplyFormatOnUsed(area, modifyFormat, formatResolver.Resolve);
    }

    private static void ApplyToAll(
        Area area,
        Func<XLCellFormatValue, XLCellFormatValue> modifyFormat,
        XLWorksheet worksheet
    )
    {
        FormatResolver formatResolver = new(worksheet);
        XLCellsCollection cellsCollection = worksheet.Internals.CellsCollection;
        cellsCollection.ApplyFormatOnAll(area, modifyFormat, formatResolver.Resolve);
    }

    /// <summary>
    /// A format value resolution hierarchy for a range API object. Each range API type needs
    /// to set proper fallbacks through ctor.
    /// </summary>
    private readonly record struct Hierarchy
    {
        private readonly XLWorkbook _workbook;
        private readonly string? _sheetName;
        private readonly int? _columnNumber;
        private readonly int? _rowNumber;
        private readonly Point? _point;

        public Hierarchy(
            XLWorkbook workbook,
            string? sheetName,
            int? columnNumber,
            int? rowNumber,
            Point? point
        )
        {
            this._workbook = workbook;
            this._sheetName = sheetName;
            this._columnNumber = columnNumber;
            this._rowNumber = rowNumber;
            this._point = point;
        }

        private XLCellFormatValue DefaultFormat => this._workbook.Styles.DefaultCellFormat;

        internal XLCellFormatValue Resolve()
        {
            bool isForWorkbook = this._sheetName is null;
            if (isForWorkbook)
            {
                return this.DefaultFormat;
            }

            // First, make sure the sheet exists
            if (!this._workbook.TryGetWorksheet(this._sheetName, out XLWorksheet sheet))
            {
                return this.DefaultFormat;
            }

            if (this._point is { } point)
            {
                FormatSlice formatSlice = sheet.Internals.CellsCollection.FormatSlice;
                XLCellFormatValue? cellFormat = formatSlice.GetFormat(point);
                if (cellFormat is not null)
                {
                    return cellFormat;
                }
            }

            if (this._rowNumber is { } rowNumber)
            {
                XLRowsCollection rowsCollection = sheet.Internals.RowsCollection;
                if (
                    rowsCollection.TryGetValue(rowNumber, out XLRow? row)
                    && row.FormatValue is { } rowFormat
                )
                {
                    return rowFormat;
                }
            }

            if (this._columnNumber is { } columnNumber)
            {
                XLColumnsCollection columnsCollection = sheet.Internals.ColumnsCollection;
                if (
                    columnsCollection.TryGetValue(columnNumber, out XLColumn? column)
                    && column.FormatValue is { } columnFormat
                )
                {
                    return columnFormat;
                }
            }

            if (sheet.FormatValue is { } sheetFormat)
            {
                return sheetFormat;
            }

            return this.DefaultFormat;
        }
    }
}
