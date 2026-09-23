using System.IO.Packaging;

namespace XlsxSharp.Tests.Utils;

internal static class ExcelDocsComparer
{
    public static bool Compare(string left, string right, out string message)
    {
        using (FileStream leftStream = File.OpenRead(left))
        using (FileStream rightStream = File.OpenRead(right))
        {
            return Compare(leftStream, rightStream, out message);
        }
    }

    public static bool Compare(Stream left, Stream right, out string message)
    {
        using (Package leftPackage = Package.Open(left, FileMode.Open, FileAccess.Read))
        using (Package rightPackage = Package.Open(right, FileMode.Open, FileAccess.Read))
        {
            return PackageHelper.Compare(
                leftPackage,
                rightPackage,
                false,
                ExcludeMethod,
                out message
            );
        }
    }

    private static bool ExcludeMethod(Uri uri)
    {
        // Relationships are compared through the parts that use them, and the core properties
        // carry the time of saving. The .psmdcp name is where System.IO.Packaging put them in
        // references written before XlsxSharp had its own packaging layer.
        if (
            uri.OriginalString.EndsWith(".rels")
            || uri.OriginalString.EndsWith(".psmdcp")
            || uri.OriginalString == "/docProps/core.xml"
        )
        {
            return true;
        }
        return false;
    }
}
