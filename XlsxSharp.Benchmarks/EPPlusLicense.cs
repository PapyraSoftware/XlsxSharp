using System.Runtime.CompilerServices;
using OfficeOpenXml;

namespace XlsxSharp.Benchmarks;

internal static class EPPlusLicense
{
    // EPPlus 8 refuses to open a package until a license is set. This benchmark project is
    // noncommercial (comparing OSS libraries), so it self-declares under the Polyform
    // Noncommercial license rather than requiring every contributor to configure one.
    [ModuleInitializer]
    public static void SetLicense() =>
        ExcelPackage.License.SetNonCommercialPersonal("XlsxSharp.Benchmarks");
}
