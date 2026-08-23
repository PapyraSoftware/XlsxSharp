namespace XlsxSharp.Tests.Excel.Drawings;

public class PictureTests
{
    [Test]
    [Arguments("Other.Drawings.picture-webp.xlsx")]
    public void CanLoadAndSaveWorkbookWithImageType(string resourceWithImageType) =>
        TestHelper.LoadSaveAndCompare(resourceWithImageType, resourceWithImageType);
}
