using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel.IO;

#nullable disable
internal class AutoFilterReader
{
    internal static void LoadAutoFilter(AutoFilter af, XLWorksheet ws)
    {
        if (af != null)
        {
            ws.Range(af.Reference.Value).SetAutoFilter();
            XLAutoFilter autoFilter = ws.AutoFilter;
            LoadAutoFilterSort(af, ws, autoFilter);
            LoadAutoFilterColumns(af, autoFilter);
        }
    }

    internal static void LoadAutoFilterColumns(AutoFilter af, XLAutoFilter autoFilter)
    {
        foreach (FilterColumn filterColumn in af.Elements<FilterColumn>())
        {
            int column = (int)filterColumn.ColumnId.Value + 1;
            XLFilterColumn xlFilterColumn = autoFilter.Column(column);
            if (filterColumn.CustomFilters is { } customFilters)
            {
                xlFilterColumn.FilterType = XLFilterType.Custom;
                XLConnector connector = OpenXmlHelper.GetBooleanValueAsBool(
                    customFilters.And,
                    false
                )
                    ? XLConnector.And
                    : XLConnector.Or;

                foreach (CustomFilter filter in customFilters.OfType<CustomFilter>())
                {
                    // Equal or NotEqual use wildcards, not value comparison. The rest does value comparison.
                    // There is no filter operation for equal of numbers (maybe combine >= and <=).
                    XLFilterOperator op = filter.Operator is not null
                        ? filter.Operator.Value.ToXlsxSharp()
                        : XLFilterOperator.Equal;
                    XLFilter xlFilter;
                    string filterValue = filter.Val.Value;
                    switch (op)
                    {
                        case XLFilterOperator.Equal:
                            xlFilter = XLFilter.CreateCustomPatternFilter(
                                filterValue,
                                true,
                                connector
                            );
                            break;
                        case XLFilterOperator.NotEqual:
                            xlFilter = XLFilter.CreateCustomPatternFilter(
                                filterValue,
                                false,
                                connector
                            );
                            break;
                        default:
                            // OOXML allows only string, so do your best to convert back to a properly typed
                            // variable. It's not perfect, but let's mimic Excel.
                            XLCellValue customValue = XLCellValue.FromText(
                                filterValue,
                                CultureInfo.InvariantCulture
                            );
                            xlFilter = XLFilter.CreateCustomFilter(customValue, op, connector);
                            break;
                    }

                    xlFilterColumn.AddFilter(xlFilter);
                }
            }
            else if (filterColumn.Filters is { } filters)
            {
                xlFilterColumn.FilterType = XLFilterType.Regular;
                foreach (Filter filter in filters.OfType<Filter>())
                {
                    xlFilterColumn.AddFilter(XLFilter.CreateRegularFilter(filter.Val.Value));
                }

                foreach (DateGroupItem dateGroupItem in filters.OfType<DateGroupItem>())
                {
                    if (
                        dateGroupItem.DateTimeGrouping is null
                        || !dateGroupItem.DateTimeGrouping.HasValue
                    )
                    {
                        continue;
                    }

                    XLDateTimeGrouping xlGrouping =
                        dateGroupItem.DateTimeGrouping.Value.ToXlsxSharp();
                    int year = 1900;
                    int month = 1;
                    int day = 1;
                    int hour = 0;
                    int minute = 0;
                    int second = 0;

                    bool valid = true;

                    if (xlGrouping >= XLDateTimeGrouping.Year)
                    {
                        if (dateGroupItem.Year?.HasValue ?? false)
                        {
                            year = dateGroupItem.Year.Value;
                        }
                        else
                        {
                            valid = false;
                        }
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Month)
                    {
                        if (dateGroupItem.Month?.HasValue ?? false)
                        {
                            month = dateGroupItem.Month.Value;
                        }
                        else
                        {
                            valid = false;
                        }
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Day)
                    {
                        if (dateGroupItem.Day?.HasValue ?? false)
                        {
                            day = dateGroupItem.Day.Value;
                        }
                        else
                        {
                            valid = false;
                        }
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Hour)
                    {
                        if (dateGroupItem.Hour?.HasValue ?? false)
                        {
                            hour = dateGroupItem.Hour.Value;
                        }
                        else
                        {
                            valid = false;
                        }
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Minute)
                    {
                        if (dateGroupItem.Minute?.HasValue ?? false)
                        {
                            minute = dateGroupItem.Minute.Value;
                        }
                        else
                        {
                            valid = false;
                        }
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Second)
                    {
                        if (dateGroupItem.Second?.HasValue ?? false)
                        {
                            second = dateGroupItem.Second.Value;
                        }
                        else
                        {
                            valid = false;
                        }
                    }

                    if (valid)
                    {
                        DateTime date = new(year, month, day, hour, minute, second);
                        XLFilter xlDateGroupFilter = XLFilter.CreateDateGroupFilter(
                            date,
                            xlGrouping
                        );
                        xlFilterColumn.AddFilter(xlDateGroupFilter);
                    }
                }
            }
            else if (filterColumn.Top10 is { } top10)
            {
                xlFilterColumn.FilterType = XLFilterType.TopBottom;
                xlFilterColumn.TopBottomType = OpenXmlHelper.GetBooleanValueAsBool(
                    top10.Percent,
                    false
                )
                    ? XLTopBottomType.Percent
                    : XLTopBottomType.Items;
                bool takeTop = OpenXmlHelper.GetBooleanValueAsBool(top10.Top, true);
                xlFilterColumn.TopBottomPart = takeTop
                    ? XLTopBottomPart.Top
                    : XLTopBottomPart.Bottom;

                // Value contains how many percent or items, so it can only be int.
                // Filter value is optional, so we don't rely on it.
                int percentsOrItems = (int)top10.Val.Value;
                xlFilterColumn.TopBottomValue = percentsOrItems;
                xlFilterColumn.AddFilter(XLFilter.CreateTopBottom(takeTop, percentsOrItems));
            }
            else if (filterColumn.DynamicFilter is { } dynamicFilter)
            {
                xlFilterColumn.FilterType = XLFilterType.Dynamic;
                XLFilterDynamicType dynamicType = dynamicFilter.Type is { } dynamicFilterType
                    ? dynamicFilterType.Value.ToXlsxSharp()
                    : XLFilterDynamicType.AboveAverage;
                double dynamicValue = filterColumn.DynamicFilter.Val.Value;

                xlFilterColumn.DynamicType = dynamicType;
                xlFilterColumn.DynamicValue = dynamicValue;
                xlFilterColumn.AddFilter(
                    XLFilter.CreateAverage(
                        dynamicValue,
                        dynamicType == XLFilterDynamicType.AboveAverage
                    )
                );
            }
        }
    }

    private static void LoadAutoFilterSort(AutoFilter af, XLWorksheet ws, XLAutoFilter autoFilter)
    {
        SortState sort = af.Elements<SortState>().FirstOrDefault();
        if (sort != null)
        {
            SortCondition condition = sort.Elements<SortCondition>().FirstOrDefault();
            if (condition != null)
            {
                int column =
                    ws.Range(condition.Reference.Value).FirstCell().Address.ColumnNumber
                    - autoFilter.Range.FirstCell().Address.ColumnNumber
                    + 1;
                autoFilter.SortColumn = column;
                autoFilter.Sorted = true;
                autoFilter.SortOrder =
                    condition.Descending != null && condition.Descending.Value
                        ? XLSortOrder.Descending
                        : XLSortOrder.Ascending;
            }
        }
    }
}
