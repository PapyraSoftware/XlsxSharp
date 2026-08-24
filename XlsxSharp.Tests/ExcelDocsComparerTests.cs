using XlsxSharp.Examples;

namespace XlsxSharp.Tests;

public class ExcelDocsComparerTests
{
    [Test]
    public void CheckEqual()
    {
        string left = ExampleHelper.GetTempFilePath("left.xlsx");
        string right = ExampleHelper.GetTempFilePath("right.xlsx");
        try
        {
            new BasicTable().Create(left);
            new BasicTable().Create(right);
            ClassicAssert.IsTrue(ExcelDocsComparer.Compare(left, right, out string _));
        }
        finally
        {
            if (File.Exists(left))
            {
                File.Delete(left);
            }
            if (File.Exists(right))
            {
                File.Delete(right);
            }
        }
    }

    [Test]
    public void CheckNonEqual()
    {
        string left = ExampleHelper.GetTempFilePath("left.xlsx");
        string right = ExampleHelper.GetTempFilePath("right.xlsx");
        try
        {
            new BasicTable().Create(left);
            new HelloWorld().Create(right);

            ClassicAssert.IsFalse(ExcelDocsComparer.Compare(left, right, out string _));
        }
        finally
        {
            if (File.Exists(left))
            {
                File.Delete(left);
            }
            if (File.Exists(right))
            {
                File.Delete(right);
            }
        }
    }
}
