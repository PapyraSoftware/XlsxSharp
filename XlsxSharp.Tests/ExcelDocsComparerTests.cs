using System.IO;
using NUnit.Framework;
using XlsxSharp.Examples;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests;

[TestFixture]
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
            Assert.IsTrue(ExcelDocsComparer.Compare(left, right, out string _));
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

            Assert.IsFalse(ExcelDocsComparer.Compare(left, right, out string _));
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
