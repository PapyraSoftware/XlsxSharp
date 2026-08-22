#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace XlsxSharp.Extensions;

internal static class GuidExtensions
{
    internal static string WrapInBraces(this Guid guid) => string.Concat('{', guid.ToString(), '}');
}
