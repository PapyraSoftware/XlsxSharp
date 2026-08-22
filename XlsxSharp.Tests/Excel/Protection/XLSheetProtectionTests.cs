using System;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Tests.Excel.Protection;

[TestFixture]
public class XlSheetProtectionTests
{
    [Test]
    public void AllowEverything()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowedElements = XLSheetProtectionElements.Everything;

            foreach (
                XLSheetProtectionElements element in Enum.GetValues(
                        typeof(XLSheetProtectionElements)
                    )
                    .Cast<XLSheetProtectionElements>()
            )
            {
                Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowElement(XLSheetProtectionElements.Everything);

            foreach (
                XLSheetProtectionElements element in Enum.GetValues(
                        typeof(XLSheetProtectionElements)
                    )
                    .Cast<XLSheetProtectionElements>()
            )
            {
                Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowEverything();

            foreach (
                XLSheetProtectionElements element in Enum.GetValues(
                        typeof(XLSheetProtectionElements)
                    )
                    .Cast<XLSheetProtectionElements>()
            )
            {
                Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }
    }

    [Test]
    public void AllowNothing()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowedElements = XLSheetProtectionElements.None;

            foreach (
                XLSheetProtectionElements element in Enum.GetValues(
                        typeof(XLSheetProtectionElements)
                    )
                    .Cast<XLSheetProtectionElements>()
                    .Where(e => e != XLSheetProtectionElements.None)
            )
            {
                Assert.IsFalse(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }

        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowNone();

            foreach (
                XLSheetProtectionElements element in Enum.GetValues(
                        typeof(XLSheetProtectionElements)
                    )
                    .Cast<XLSheetProtectionElements>()
                    .Where(e => e != XLSheetProtectionElements.None)
            )
            {
                Assert.IsFalse(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }
    }

    [Test]
    public void ChangeHashingAlgorithm()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.AddWorksheet();
                ws.Protect("123");

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                Assert.IsTrue(ws.Protection.IsProtected);
                Assert.AreEqual(Algorithm.SimpleHash, ws.Protection.Algorithm);

                ws.Unprotect("123");
                ws.Protect("123", Algorithm.SHA512);
                wb.Save();
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                Assert.IsTrue(ws.Protection.IsProtected);
                Assert.AreEqual(Algorithm.SHA512, ws.Protection.Algorithm);

                Assert.DoesNotThrow(() => ws.Unprotect("123"));
            }
        }
    }

    [Test]
    public void CopyProtectionFromAnotherSheet()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Misc\SheetProtection.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws1 = wb.Worksheet("Protected Password = 123");
            XLSheetProtection p1 = ws1.Protection.CastTo<XLSheetProtection>();
            Assert.IsTrue(p1.IsProtected);

            IXLWorksheet ws2 = ws1.CopyTo("New worksheet");
            Assert.IsFalse(ws2.Protection.IsProtected);
            XLSheetProtection p2 = ws2.Protection.CopyFrom(p1).CastTo<XLSheetProtection>();

            Assert.IsTrue(p2.IsProtected);
            Assert.IsTrue(p2.IsPasswordProtected);
            Assert.AreEqual(p1.Algorithm, p2.Algorithm);
            Assert.AreEqual(p1.PasswordHash, p2.PasswordHash);
            Assert.AreEqual(p1.Base64EncodedSalt, p2.Base64EncodedSalt);
            Assert.AreEqual(p1.SpinCount, p2.SpinCount);

            Assert.IsTrue(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertColumns));
            Assert.IsTrue(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertRows));
            Assert.IsFalse(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertHyperlinks));

            Assert.Throws<InvalidOperationException>(() => ws2.Unprotect());
            ws2.Unprotect("123");
        }
    }

    [Test]
    public void SetWorksheetProtectionCloning()
    {
        IXLWorksheet ws1 = new XLWorkbook().AddWorksheet();
        IXLWorksheet ws2 = new XLWorkbook().AddWorksheet();

        ws1.Protect("123")
            .AllowElement(XLSheetProtectionElements.FormatEverything)
            .DisallowElement(XLSheetProtectionElements.FormatCells);

        Assert.AreEqual(
            XLSheetProtectionElements.FormatColumns
                | XLSheetProtectionElements.FormatRows
                | XLSheetProtectionElements.SelectEverything,
            ws1.Protection.AllowedElements
        );

        ws2.Protection = ws1.Protection;

        Assert.IsFalse(ReferenceEquals(ws1.Protection, ws2.Protection));
        Assert.IsTrue(ws2.Protection.IsProtected);
        Assert.AreEqual(
            XLSheetProtectionElements.FormatColumns
                | XLSheetProtectionElements.FormatRows
                | XLSheetProtectionElements.SelectEverything,
            ws2.Protection.AllowedElements
        );
        Assert.AreEqual(
            (ws1.Protection as XLSheetProtection).PasswordHash,
            (ws2.Protection as XLSheetProtection).PasswordHash
        );
    }

    [Test]
    public void TestUnprotectWorksheetWithNoPassword()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet("Sheet1");
            Assert.IsTrue(ws.Protection.IsProtected);
            ws.Unprotect();
            Assert.IsFalse(ws.Protection.IsProtected);
        }
    }

    [Test]
    public void TestWorksheetWithSha512Protection()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet("Sheet2");
            Assert.IsTrue(ws.Protection.IsProtected);

            // Password required
            Assert.Throws<InvalidOperationException>(() => ws.Unprotect());

            Assert.AreEqual(Algorithm.SHA512, ws.Protection.Algorithm);
            ws.Unprotect("abc");
            Assert.IsFalse(ws.Protection.IsProtected);
        }
    }
}
