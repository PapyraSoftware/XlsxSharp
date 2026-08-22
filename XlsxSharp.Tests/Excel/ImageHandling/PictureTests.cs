using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using NUnit.Framework;
using XlsxSharp.Examples;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Drawings;
using Point = System.Drawing.Point;

namespace XlsxSharp.Tests.Excel.ImageHandling;

[TestFixture]
public class PictureTests
{
    [Test]
    public void CanAddPictureFromStream()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            using (
                Stream? resourceStream = Assembly
                    .GetAssembly(typeof(BasicTable))
                    .GetManifestResourceStream("XlsxSharp.Examples.Resources.SampleImage.jpg")
            )
            {
                IXLPicture picture = ws.AddPicture(resourceStream, "MyPicture")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 50)
                    .WithSize(200, 200);

                Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(200, picture.Width);
                Assert.AreEqual(200, picture.Height);
            }
        }
    }

    [Test]
    public void CanAddPictureFromFile()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            string path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

            try
            {
                using (
                    Stream? resourceStream = Assembly
                        .GetAssembly(typeof(BasicTable))
                        .GetManifestResourceStream("XlsxSharp.Examples.Resources.SampleImage.jpg")
                )
                using (FileStream fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                IXLPicture picture = ws.AddPicture(path)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 50);

                Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(400, picture.Width);
                Assert.AreEqual(400, picture.Height);
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }
    }

    [Test]
    public void CanAddPictureConcurrentlyFromFile()
    {
        string path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

        try
        {
            using (
                Stream? resourceStream = Assembly
                    .GetAssembly(typeof(BasicTable))
                    .GetManifestResourceStream("XlsxSharp.Examples.Resources.SampleImage.jpg")
            )
            using (FileStream fileStream = File.Create(path))
            {
                resourceStream.Seek(0, SeekOrigin.Begin);
                resourceStream.CopyTo(fileStream);
                fileStream.Close();
            }

            Parallel.Invoke(() => verifyAddImageFromFile(path), () => verifyAddImageFromFile(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    private static void verifyAddImageFromFile(string filePath)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            IXLPicture picture = ws.AddPicture(filePath)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50);

            Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
            Assert.AreEqual(400, picture.Width);
            Assert.AreEqual(50, picture.Top);
        }
    }

    [Test]
    public void CanScaleImage()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            using (
                Stream? resourceStream = Assembly
                    .GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
            )
            {
                IXLPicture pic = ws.AddPicture(resourceStream, "MyPicture")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 50);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(252, pic.Width);
                Assert.AreEqual(152, pic.Height);

                pic.ScaleHeight(0.7);
                pic.ScaleWidth(1.2);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(302, pic.Width);
                Assert.AreEqual(106, pic.Height);

                pic.ScaleHeight(0.7);
                pic.ScaleWidth(1.2);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(362, pic.Width);
                Assert.AreEqual(74, pic.Height);

                pic.ScaleHeight(0.8, true);
                pic.ScaleWidth(1.1, true);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(277, pic.Width);
                Assert.AreEqual(122, pic.Height);
            }
        }
    }

    [Test]
    public void TestDefaultPictureNames()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            using (
                Stream? stream = Assembly
                    .GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
            )
            {
                ws.AddPicture(stream, XLPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XLPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
                stream.Position = 0;

                ws.AddPicture(stream, XLPictureFormat.Png);
                stream.Position = 0;
            }

            Assert.AreEqual("Picture 1", ws.Pictures.Skip(0).First().Name);
            Assert.AreEqual("Picture 2", ws.Pictures.Skip(1).First().Name);
            Assert.AreEqual("Picture 4", ws.Pictures.Skip(2).First().Name);
            Assert.AreEqual("Picture 5", ws.Pictures.Skip(3).First().Name);
        }
    }

    [Test]
    public void TestDefaultIds()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            using (
                Stream? stream = Assembly
                    .GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
            )
            {
                ws.AddPicture(stream, XLPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XLPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
                stream.Position = 0;

                ws.AddPicture(stream, XLPictureFormat.Png);
                stream.Position = 0;
            }

            Assert.AreEqual(1, ws.Pictures.Skip(0).First().Id);
            Assert.AreEqual(2, ws.Pictures.Skip(1).First().Id);
            Assert.AreEqual(3, ws.Pictures.Skip(2).First().Id);
            Assert.AreEqual(4, ws.Pictures.Skip(3).First().Id);
        }
    }

    [Test]
    public void XLMarkerTests()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        XLMarker firstMarker = new(ws.Cell(1, 10), new Point(100, 0));

        Assert.AreEqual(10, firstMarker.ColumnNumber);
        Assert.AreEqual(1, firstMarker.RowNumber);
        Assert.AreEqual(100, firstMarker.Offset.X);
        Assert.AreEqual(0, firstMarker.Offset.Y);
    }

    [Test]
    public void XLPictureTests()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            using (
                Stream? stream = Assembly
                    .GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
            )
            {
                IXLPicture pic = ws.AddPicture(stream, XLPictureFormat.Png, "Image1")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155);

                Assert.AreEqual(XLPicturePlacement.FreeFloating, pic.Placement);
                Assert.AreEqual("Image1", pic.Name);
                Assert.AreEqual(XLPictureFormat.Png, pic.Format);
                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(252, pic.Width);
                Assert.AreEqual(152, pic.Height);
                Assert.AreEqual(220, pic.Left);
                Assert.AreEqual(155, pic.Top);
            }
        }
    }

    [Test]
    public void CanLoadFileWithImagesAndCopyImagesToNewSheet()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            Assert.AreEqual(2, ws.Pictures.Count);

            IXLWorksheet copy = ws.CopyTo("NewSheet");
            Assert.AreEqual(2, copy.Pictures.Count);
        }
    }

    [Test]
    public void CanDeletePictures()
    {
        using (MemoryStream ms = new())
        {
            int originalCount;

            using (
                Stream stream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")
                )
            )
            using (XLWorkbook wb = new(stream))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                originalCount = ws.Pictures.Count;
                ws.Pictures.Delete(ws.Pictures.First());

                string pictureName = ws.Pictures.First().Name;
                ws.Pictures.Delete(pictureName);

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                Assert.AreEqual(originalCount - 2, ws.Pictures.Count);
            }
        }
    }

    [Test]
    public void PictureRenameTests()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet("Images3");
            IXLPicture picture = ws.Pictures.First();
            Assert.AreEqual("Picture 1", picture.Name);

            picture.Name = "picture 1";
            picture.Name = "pICture 1";
            picture.Name = "Picture 1";

            picture = ws.Pictures.Last();
            picture.Name = "new name";

            Assert.Throws<ArgumentException>(() => picture.Name = "Picture 1");
            Assert.Throws<ArgumentException>(() => picture.Name = "picTURE 1");
        }
    }

    [Test]
    public void HandleDuplicatePictureIdsAcrossWorksheets()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet("Sheet1");
            IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");

            using (
                Stream? stream = Assembly
                    .GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
            )
            {
                (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2);
                (ws1 as XLWorksheet).AddPicture(stream, "Picture 2", 3);

                //Internal method - used for loading files
                XLPicture? pic =
                    (ws2 as XLWorksheet)
                        .AddPicture(stream, "Picture 1", 2)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(220, 155) as XLPicture;

                int id = pic.Id;

                pic.Id = id;
                Assert.AreEqual(id, pic.Id);

                pic.Id = 3;
                Assert.AreEqual(3, pic.Id);

                pic.Id = id;

                XLPicture? pic2 =
                    (ws2 as XLWorksheet)
                        .AddPicture(stream, "Picture 2", 3)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(440, 300) as XLPicture;
            }
        }
    }

    [Test]
    public void CopyImageSameWorksheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");

        IXLPicture original;
        using (
            Stream? stream = Assembly
                .GetExecutingAssembly()
                .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
        )
        {
            original =
                (ws1 as XLWorksheet)
                    .AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XLPicture;
        }

        XLPicture? copy = original.Duplicate().MoveTo(300, 200) as XLPicture;

        Assert.AreEqual(2, ws1.Pictures.Count);
        Assert.AreEqual(ws1, copy.Worksheet);
        Assert.AreEqual(original.Format, copy.Format);
        Assert.AreEqual(original.Height, copy.Height);
        Assert.AreEqual(original.Placement, copy.Placement);
        Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
        Assert.AreEqual(original.Width, copy.Width);
        Assert.AreEqual(
            original.ImageStream.ToArray(),
            copy.ImageStream.ToArray(),
            "Image streams differ"
        );

        Assert.AreEqual(200, copy.Top);
        Assert.AreEqual(300, copy.Left);
        Assert.AreNotEqual(original.Id, copy.Id);
        Assert.AreNotEqual(original.Name, copy.Name);
    }

    [Test]
    public void CopyImageDifferentWorksheets()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        IXLPicture original;
        using (
            Stream? stream = Assembly
                .GetExecutingAssembly()
                .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
        )
        {
            original =
                (ws1 as XLWorksheet)
                    .AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XLPicture;
        }
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");

        IXLPicture copy = original.CopyTo(ws2);

        Assert.AreEqual(1, ws1.Pictures.Count);
        Assert.AreEqual(1, ws2.Pictures.Count);

        Assert.AreEqual(ws2, copy.Worksheet);

        Assert.AreEqual(original.Format, copy.Format);
        Assert.AreEqual(original.Height, copy.Height);
        Assert.AreEqual(original.Left, copy.Left);
        Assert.AreEqual(original.Name, copy.Name);
        Assert.AreEqual(original.Placement, copy.Placement);
        Assert.AreEqual(original.Top, copy.Top);
        Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
        Assert.AreEqual(original.Width, copy.Width);
        Assert.AreEqual(
            original.ImageStream.ToArray(),
            copy.ImageStream.ToArray(),
            "Image streams differ"
        );

        Assert.AreNotEqual(original.Id, copy.Id);
    }

    [Test]
    public void PictureShiftsWhenInsertingRows()
    {
        using (XLWorkbook wb = new())
        using (
            Stream? stream = Assembly
                .GetExecutingAssembly()
                .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")
        )
        {
            IXLWorksheet ws = wb.Worksheets.Add("ImageShift");
            IXLPicture picture = ws.AddPicture(stream, XLPictureFormat.Png, "PngImage")
                .MoveTo(ws.Cell(5, 2))
                .WithPlacement(XLPicturePlacement.Move);

            ws.Row(2).InsertRowsBelow(20);

            Assert.AreEqual(25, picture.TopLeftCell.Address.RowNumber);
        }
    }

    [Test]
    public void PictureNotFound()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Picture("dummy"));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Pictures.Delete("dummy"));
        }
    }

    [Test]
    public void CanCopyEmfPicture()
    {
        // #1621 - There are 2 Bmp Guids: ImageFormat.Bmp and ImageFormat.MemoryBmp
        using Stream stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\Pictures\EmfPicture.xlsx")
        );
        using XLWorkbook wb = new(stream);
        IXLWorksheet ws1 = wb.Worksheets.First();
        IXLPicture img1 = ws1.Pictures.First();

        IXLWorksheet ws2 = wb.AddWorksheet();

        IXLPicture img2 = img1.CopyTo(ws2);

        Assert.AreEqual(XLPictureFormat.Emf, img2.Format);

        using MemoryStream ms = new();
        wb.SaveAs(ms);

        ms.Seek(0, SeekOrigin.Begin);

        using XLWorkbook wb2 = new(ms);
        ws2 = wb2.Worksheet("Sheet2");
        img2 = ws2.Pictures.First();
        Assert.AreEqual(XLPictureFormat.Emf, img2.Format);
    }

    [Test]
    public void KeepOriginalDrawingShapesZOrder() =>
        // File contains shapes and a picture in a mixed order.
        TestHelper.LoadSaveAndCompare(
            @"Other\Pictures\ImageShapeZOrder-Input.xlsx",
            @"Other\Pictures\ImageShapeZOrder-Output.xlsx"
        );

    [TestCase("@")]
    [TestCase(":")]
    [TestCase("\\")]
    [TestCase("/")]
    [TestCase("?")]
    [TestCase("*")]
    [TestCase("[]")]
    [TestCase(" ")] // Whitespace name is allowed, but can't be empty
    [TestCase("C:\\Images\\pic.jpg")] // Path with multiple forbidden chars
    [TestCase("http://example.com/image.jpg")] // URL with multiple forbidden chars
    [TestCase("Picture@01\\QPosted@")] // A name from a problematic workbook
    public void PictureCanHaveUnusualCharactersInName(string nameWithUnusualCharacter) =>
        // The name of a picture couldn't contain certain characters in some ancient version of Excel. Verify that
        // it is no longer the case through the whole lifecycle (add picture, change name, save, load).
        AssertPictureNameAllowed(nameWithUnusualCharacter);

    [Test]
    public void PictureNameCanBeLong()
    {
        // Picture name was originally limited to 31 characters. Verify that it is no longer the case.
        string longName = new('a', 100_000);
        AssertPictureNameAllowed(longName);
    }

    [TestCase(null)]
    [TestCase("")]
    public void PictureNameCantBeNullOrEmpty(string invalidName)
    {
        // Picture name is a required attribute, though Excel generates a name if it isn't specified instead of failing.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        using Stream? imageStream = Assembly
            .GetExecutingAssembly()
            .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png");
        IXLPicture picture = ws.AddPicture(imageStream);
        Assert.Throws<ArgumentException>(() => picture.Name = invalidName);
    }

    private static void AssertPictureNameAllowed(string testedName) =>
        TestHelper.CreateSaveLoadAssert(
            wb =>
            {
                using Stream? imageStream = Assembly
                    .GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png");
                IXLWorksheet ws1 = wb.AddWorksheet("AddPicture");
                ws1.AddPicture(imageStream, testedName);

                IXLWorksheet ws2 = wb.AddWorksheet("Setter");
                IXLPicture picture = ws2.AddPicture(imageStream);
                picture.Name = testedName;
            },
            wb =>
            {
                Assert.AreEqual(testedName, wb.Worksheet("AddPicture").Pictures.Single().Name);
                Assert.AreEqual(testedName, wb.Worksheet("Setter").Pictures.Single().Name);
            }
        );
}
