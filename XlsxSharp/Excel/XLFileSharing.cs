// Keep this file CodeMaid organised and cleaned

namespace XlsxSharp.Excel;

internal class XLFileSharing : IXLFileSharing
{
    public bool ReadOnlyRecommended { get; set; }
    public string? UserName { get; set; }
}
