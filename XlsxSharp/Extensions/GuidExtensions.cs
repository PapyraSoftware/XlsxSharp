#nullable disable

using System;

namespace XlsxSharp.Extensions;

internal static class GuidExtensions
{
    internal static string WrapInBraces(this Guid guid) => guid.ToString("B");
}
