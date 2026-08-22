using NUnit.Framework;

namespace XlsxSharp.Tests.Excel.Drawings;

[TestFixture]
public class PictureTests
{
    [TestCase("Other.Drawings.picture-webp.xlsx")]
    public void CanLoadAndSaveWorkbookWithImageType(string resourceWithImageType)
    {
        TestHelper.LoadSaveAndCompare(resourceWithImageType, resourceWithImageType);
    }
}
