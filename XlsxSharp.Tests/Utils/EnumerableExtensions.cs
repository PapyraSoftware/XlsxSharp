using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Utils;

internal static class EnumerableExtensions
{
    public static string ToSpaceList(
        this IEnumerable<IXLRange> ranges,
        bool includeSheet = false
    ) =>
        string.Join(
            " ",
            ranges.Select(r => r.RangeAddress.ToString(XLReferenceStyle.A1, includeSheet))
        );

    public static string ToSpaceList(this IEnumerable<Area> areas) => string.Join(" ", areas);
}
