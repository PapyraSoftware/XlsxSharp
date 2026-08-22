#nullable disable

// Keep this file CodeMaid organised and cleaned
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel.Protection;

public interface IXLSheetProtection : IXLElementProtection<XLSheetProtectionElements>
{
    IXLSheetProtection Protect(XLSheetProtectionElements allowedElements);

    IXLSheetProtection Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements);

    IXLSheetProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm,
        XLSheetProtectionElements allowedElements = XLSheetProtectionElements.SelectEverything
    );
}
