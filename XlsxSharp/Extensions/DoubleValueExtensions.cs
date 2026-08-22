#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using DocumentFormat.OpenXml;

namespace XlsxSharp.Extensions;

internal static class DoubleValueExtensions
{
    public static DoubleValue SaveRound(this DoubleValue value) =>
        value.HasValue ? new DoubleValue(Math.Round(value, 6)) : value;
}
