#nullable disable

// Keep this file CodeMaid organised and cleaned
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel.Protection;

public interface IXLSheetProtection : IXLElementProtection<XLSheetProtectionElements>
{
    public IXLSheetProtection Protect(XLSheetProtectionElements allowedElements);

    public IXLSheetProtection Protect(
        Algorithm algorithm,
        XLSheetProtectionElements allowedElements
    );

    public IXLSheetProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm,
        XLSheetProtectionElements allowedElements = XLSheetProtectionElements.SelectEverything
    );
}
