#nullable disable

using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel.Protection;

public interface IXLWorkbookProtection : IXLElementProtection<XLWorkbookProtectionElements>
{
    public IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements);

    public IXLWorkbookProtection Protect(
        Algorithm algorithm,
        XLWorkbookProtectionElements allowedElements
    );

    public IXLWorkbookProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm,
        XLWorkbookProtectionElements allowedElements = XLWorkbookProtectionElements.Windows
    );
}
