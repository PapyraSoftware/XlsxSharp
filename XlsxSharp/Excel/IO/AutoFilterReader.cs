using System.Globalization;
using System.Xml.Linq;
using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel.IO;

#nullable disable
internal class AutoFilterReader
{
    internal static void LoadAutoFilter(XElement af, XLWorksheet ws)
    {
        if (af is not null)
        {
            ws.Range(SpreadsheetXml.String(af, "ref")).SetAutoFilter();
            XLAutoFilter autoFilter = ws.AutoFilter;
            LoadAutoFilterSort(af, ws, autoFilter);
            LoadAutoFilterColumns(af, autoFilter);
        }
    }

    internal static void LoadAutoFilterColumns(XElement af, XLAutoFilter autoFilter)
    {
        foreach (XElement filterColumn in af.Elements(SpreadsheetXml.Main + "filterColumn"))
        {
            int column = checked((int)SpreadsheetXml.UInt(filterColumn, "colId")) + 1;
            XLFilterColumn xlFilterColumn = autoFilter.Column(column);
            if (filterColumn.Element(SpreadsheetXml.Main + "customFilters") is { } customFilters)
            {
                LoadCustomFilters(customFilters, xlFilterColumn);
            }
            else if (filterColumn.Element(SpreadsheetXml.Main + "filters") is { } filters)
            {
                LoadRegularFilters(filters, xlFilterColumn);
            }
            else if (filterColumn.Element(SpreadsheetXml.Main + "top10") is { } top10)
            {
                LoadTopBottomFilter(top10, xlFilterColumn);
            }
            else if (
                filterColumn.Element(SpreadsheetXml.Main + "dynamicFilter") is { } dynamicFilter
            )
            {
                LoadDynamicFilter(dynamicFilter, xlFilterColumn);
            }
        }
    }

    private static void LoadCustomFilters(XElement customFilters, XLFilterColumn xlFilterColumn)
    {
        xlFilterColumn.FilterType = XLFilterType.Custom;
        XLConnector connector =
            SpreadsheetXml.Bool(customFilters, "and") ?? false ? XLConnector.And : XLConnector.Or;

        foreach (XElement filter in customFilters.Elements(SpreadsheetXml.Main + "customFilter"))
        {
            // Equal or NotEqual use wildcards, not value comparison. The rest does value comparison.
            // There is no filter operation for equal of numbers (maybe combine >= and <=).
            XLFilterOperator op = SpreadsheetXml.String(filter, "operator") is { } operatorName
                ? WorksheetXmlEnums.ParseFilterOperator(operatorName)
                : XLFilterOperator.Equal;
            string filterValue = SpreadsheetXml.String(filter, "val");
            XLFilter xlFilter = op switch
            {
                XLFilterOperator.Equal => XLFilter.CreateCustomPatternFilter(
                    filterValue,
                    true,
                    connector
                ),
                XLFilterOperator.NotEqual => XLFilter.CreateCustomPatternFilter(
                    filterValue,
                    false,
                    connector
                ),
                // OOXML allows only string, so do your best to convert back to a properly typed
                // variable. It's not perfect, but let's mimic Excel.
                _ => XLFilter.CreateCustomFilter(
                    XLCellValue.FromText(filterValue, CultureInfo.InvariantCulture),
                    op,
                    connector
                ),
            };

            xlFilterColumn.AddFilter(xlFilter);
        }
    }

    private static void LoadRegularFilters(XElement filters, XLFilterColumn xlFilterColumn)
    {
        xlFilterColumn.FilterType = XLFilterType.Regular;
        foreach (XElement filter in filters.Elements(SpreadsheetXml.Main + "filter"))
        {
            xlFilterColumn.AddFilter(
                XLFilter.CreateRegularFilter(SpreadsheetXml.String(filter, "val"))
            );
        }

        foreach (XElement dateGroupItem in filters.Elements(SpreadsheetXml.Main + "dateGroupItem"))
        {
            if (SpreadsheetXml.String(dateGroupItem, "dateTimeGrouping") is not { } grouping)
            {
                continue;
            }

            if (
                ReadDateGroup(dateGroupItem, WorksheetXmlEnums.ParseDateTimeGrouping(grouping)) is
                { } filter
            )
            {
                xlFilterColumn.AddFilter(filter);
            }
        }
    }

    /// <summary>
    /// A date group filter names the date only down to the unit it groups by, so a grouping by
    /// month carries a year and a month and nothing else. Every part down to that unit has to be
    /// there - one that is missing leaves a date that cannot be built, and the item is dropped.
    /// </summary>
    private static XLFilter ReadDateGroup(XElement dateGroupItem, XLDateTimeGrouping xlGrouping)
    {
        int year = 1900;
        int month = 1;
        int day = 1;
        int hour = 0;
        int minute = 0;
        int second = 0;

        bool valid =
            Part(XLDateTimeGrouping.Year, "year", ref year)
            && Part(XLDateTimeGrouping.Month, "month", ref month)
            && Part(XLDateTimeGrouping.Day, "day", ref day)
            && Part(XLDateTimeGrouping.Hour, "hour", ref hour)
            && Part(XLDateTimeGrouping.Minute, "minute", ref minute)
            && Part(XLDateTimeGrouping.Second, "second", ref second);

        return valid
            ? XLFilter.CreateDateGroupFilter(
                new DateTime(year, month, day, hour, minute, second),
                xlGrouping
            )
            : null;

        bool Part(XLDateTimeGrouping unit, string attributeName, ref int value)
        {
            if (xlGrouping < unit)
            {
                return true;
            }

            if (SpreadsheetXml.UInt(dateGroupItem, attributeName) is not { } read)
            {
                return false;
            }

            value = checked((int)read);
            return true;
        }
    }

    private static void LoadTopBottomFilter(XElement top10, XLFilterColumn xlFilterColumn)
    {
        xlFilterColumn.FilterType = XLFilterType.TopBottom;
        xlFilterColumn.TopBottomType =
            SpreadsheetXml.Bool(top10, "percent") ?? false
                ? XLTopBottomType.Percent
                : XLTopBottomType.Items;
        bool takeTop = SpreadsheetXml.Bool(top10, "top") ?? true;
        xlFilterColumn.TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

        // Value contains how many percent or items, so it can only be int.
        // Filter value is optional, so we don't rely on it.
        int percentsOrItems = (int)SpreadsheetXml.Double(top10, "val").Value;
        xlFilterColumn.TopBottomValue = percentsOrItems;
        xlFilterColumn.AddFilter(XLFilter.CreateTopBottom(takeTop, percentsOrItems));
    }

    private static void LoadDynamicFilter(XElement dynamicFilter, XLFilterColumn xlFilterColumn)
    {
        xlFilterColumn.FilterType = XLFilterType.Dynamic;
        XLFilterDynamicType dynamicType = SpreadsheetXml.String(dynamicFilter, "type")
            is { } dynamicFilterType
            ? WorksheetXmlEnums.ParseFilterDynamicType(dynamicFilterType)
            : XLFilterDynamicType.AboveAverage;
        double dynamicValue = SpreadsheetXml.Double(dynamicFilter, "val").Value;

        xlFilterColumn.DynamicType = dynamicType;
        xlFilterColumn.DynamicValue = dynamicValue;
        xlFilterColumn.AddFilter(
            XLFilter.CreateAverage(dynamicValue, dynamicType == XLFilterDynamicType.AboveAverage)
        );
    }

    private static void LoadAutoFilterSort(XElement af, XLWorksheet ws, XLAutoFilter autoFilter)
    {
        XElement condition = af.Element(SpreadsheetXml.Main + "sortState")
            ?.Element(SpreadsheetXml.Main + "sortCondition");
        if (condition is null)
        {
            return;
        }

        int column =
            ws.Range(SpreadsheetXml.String(condition, "ref")).FirstCell().Address.ColumnNumber
            - autoFilter.Range.FirstCell().Address.ColumnNumber
            + 1;
        autoFilter.SortColumn = column;
        autoFilter.Sorted = true;
        autoFilter.SortOrder =
            SpreadsheetXml.Bool(condition, "descending") ?? false
                ? XLSortOrder.Descending
                : XLSortOrder.Ascending;
    }
}
