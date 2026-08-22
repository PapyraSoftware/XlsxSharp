using System;

namespace XlsxSharp.Excel.Sort;

internal record XLSortElement(
    int ElementNumber,
    XLSortOrder SortOrder,
    bool IgnoreBlanks,
    bool MatchCase
) : IXLSortElement;
