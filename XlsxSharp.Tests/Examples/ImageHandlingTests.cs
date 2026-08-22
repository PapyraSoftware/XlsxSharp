using NUnit.Framework;
using XlsxSharp.Examples.ImageHandling;

namespace XlsxSharp.Tests.Examples;

[TestFixture]
public class ImageHandlingTests
{
    [Test]
    public void ImageAnchors() =>
        TestHelper.RunTestExample<ImageAnchors>(@"ImageHandling\ImageAnchors.xlsx");

    [Test]
    public void ImageFormats() =>
        TestHelper.RunTestExample<ImageFormats>(@"ImageHandling\ImageFormats.xlsx");
}
