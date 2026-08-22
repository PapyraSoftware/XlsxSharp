using NUnit.Framework;
using XlsxSharp.Utils;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Tests.Excel.Protection;

public class HashAlgorithmTests
{
    [Test]
    public void TestEmptyPassword()
    {
        Assert.IsEmpty(CryptographicAlgorithms.GetPasswordHash(Algorithm.SHA512, string.Empty));
        Assert.IsEmpty(CryptographicAlgorithms.GetPasswordHash(Algorithm.SimpleHash, string.Empty));
    }

    [Test]
    public void TestSha512()
    {
        string hash = CryptographicAlgorithms.GetPasswordHash(
            Algorithm.SHA512,
            "12345",
            "aVvPw1DNH3evPqRAd/y3UQ==",
            100000
        );
        Assert.AreEqual(
            "E+qAhyIg/HM0dUrPaENfimFOZp7wlOkJsf/sdG+AGHOA9grOv7VLb1ik2vuYohljI9G36e0ea9wnixCK0MMuyQ==",
            hash
        );
    }

    [Test]
    public void TestSimple()
    {
        string hash = CryptographicAlgorithms.GetPasswordHash(Algorithm.SimpleHash, "12345");
        Assert.AreEqual("CA9C", hash);
    }
}
