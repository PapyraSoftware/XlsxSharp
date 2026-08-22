using System;

namespace XlsxSharp.Excel.Sort;

internal record XLSortElement(
    Int32 ElementNumber,
    XLSortOrder SortOrder,
    Boolean IgnoreBlanks,
    Boolean MatchCase
) : IXLSortElement;
