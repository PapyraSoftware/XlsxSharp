namespace XlsxSharp.Excel.Formatting;

internal record XLDifferentialProtectionValue
{
    internal static readonly XLDifferentialProtectionValue Empty = new()
    {
        Hidden = null,
        Locked = null,
    };

    internal required bool? Hidden { get; init; }

    internal required bool? Locked { get; init; }
}
