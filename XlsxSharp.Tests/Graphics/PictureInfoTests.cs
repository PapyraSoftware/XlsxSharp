using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Graphics;
using Assembly = System.Reflection.Assembly;

namespace XlsxSharp.Tests.Graphics;

public class PictureInfoTests
{
    [Test]
    public async Task CanReadPng() =>
        await AssertRasterImage("SampleImagePng.png", XLPictureFormat.Png, new Size(252, 152), 96, 96);

    [Test]
    [Arguments("SampleImageJfif.jpg", 176, 270, 96, 96)]
    [Arguments("jpeg-rgb.jpg", 200, 200, 0, 0)] // Adobe JPG, has APP14 marker right after SOI instead of APP0
    public async Task CanReadJfif(string filename, int widthPx, int heightPx, int dpiX, int dpiY) =>
        await AssertRasterImage(
            $"Jpg.{filename}",
            XLPictureFormat.Jpeg,
            new Size(widthPx, heightPx),
            dpiX,
            dpiY
        );

    [Test]
    public async Task CanReadExif() =>
        await AssertRasterImage("SampleImageExif.jpg", XLPictureFormat.Jpeg, new Size(252, 152), 0, 0);

    [Test]
    public async Task CanReadGif87Image() =>
        await AssertRasterImage("SampleImageGif87a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);

    [Test]
    public async Task CanReadGif89Image() =>
        await AssertRasterImage("SampleImageGif89a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);

    [Test]
    [Arguments("SampleImageBmpWin24bit.bmp")]
    [Arguments("SampleImageBmpWin8bit.bmp")]
    [Arguments("SampleImageBmpWin4bit.bmp")]
    [Arguments("SampleImageBmpWin24bit.bmp")]
    public async Task CanReadBmpImageV3AndFurther(string imageName) =>
        await AssertRasterImage(imageName, XLPictureFormat.Bmp, new Size(167, 51), 80.645d, 80.645d);

    [Test]
    public async Task CanReadBmpV1() =>
        await AssertRasterImage("SampleImageBmpV1.bmp", XLPictureFormat.Bmp, new Size(150, 50), 0, 0);

    [Test]
    public async Task CanReadTiffWithBigEndianEncoding() =>
        await AssertRasterImage(
            "SampleImageTiffBigEndian.tiff",
            XLPictureFormat.Tiff,
            new Size(130, 45),
            96,
            96
        );

    [Test]
    public async Task CanReadTiffWithLittleEndianEncoding() =>
        await AssertRasterImage(
            "SampleImageTiffLittleEndian.tiff",
            XLPictureFormat.Tiff,
            new Size(130, 45),
            96,
            96
        );

    [Test]
    public async Task CanReadPcx() =>
        await AssertRasterImage("SampleImagePcx.pcx", XLPictureFormat.Pcx, new Size(100, 50), 96, 96);

    [Test]
    public async Task CanReadWmfWithPlaceableHeader() =>
        await AssertVectorImage("SampleImagePlaceableWmf.wmf", XLPictureFormat.Wmf, new Size(1000, 500));

    [Test]
    public async Task CanReadWmfWithOriginalHeader() =>
        await AssertVectorImage("SampleImageOriginalWmf.wmf", XLPictureFormat.Wmf, new Size(12496, 6247));

    [Test]
    public async Task CanReadEmf() =>
        await AssertVectorImage("SampleImageEmf.emf", XLPictureFormat.Emf, new Size(28844, 28938));

    [Test]
    public async Task CanReadExtendedWebp() =>
        await AssertRasterImage(
            "SampleImageWebpExtendedFormat.webp",
            XLPictureFormat.Webp,
            new Size(188, 231),
            72,
            72
        );

    [Test]
    public async Task CanReadLossyWebp() =>
        await AssertRasterImage(
            "SampleImageWebpLossy.webp",
            XLPictureFormat.Webp,
            new Size(278, 90),
            72,
            72
        );

    [Test]
    public async Task CanReadLosslessWebp() =>
        await AssertRasterImage(
            "SampleImageWebpLossless.webp",
            XLPictureFormat.Webp,
            new Size(395, 136),
            72,
            72
        );

    private static async Task AssertRasterImage(
        string imageName,
        XLPictureFormat expectedFormat,
        Size expectedPxSize,
        double expectedDpiX,
        double expectedDpiY
    ) =>
        await AssertImage(
            imageName,
            expectedFormat,
            expectedPxSize,
            Size.Empty,
            expectedDpiX,
            expectedDpiY
        );

    private static async Task AssertVectorImage(
        string imageName,
        XLPictureFormat expectedFormat,
        Size expectedHiMetricSize
    ) => await AssertImage(imageName, expectedFormat, Size.Empty, expectedHiMetricSize, 0, 0);

    private static async Task AssertImage(
        string imageName,
        XLPictureFormat expectedFormat,
        Size expectedPxSize,
        Size expectedHiMetricSize,
        double expectedDpiX,
        double expectedDpiY
    )
    {
        using Stream? stream = Assembly
            .GetExecutingAssembly()
            .GetManifestResourceStream($"XlsxSharp.Tests.Resource.Images.{imageName}");
        XLPictureInfo info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(
            stream,
            XLPictureFormat.Unknown
        );

        await Assert.That(info.Format).IsEqualTo(expectedFormat);
        await Assert.That(info.SizePx).IsEqualTo(expectedPxSize);
        await Assert.That(info.SizePhys).IsEqualTo(expectedHiMetricSize);

        // Some DPI is stored as pixels per meter, causing a rounding errors.
        await Assert.That(info.DpiX).IsCloseTo(expectedDpiX, 0.02);
        await Assert.That(info.DpiY).IsCloseTo(expectedDpiY, 0.02);
    }
}
