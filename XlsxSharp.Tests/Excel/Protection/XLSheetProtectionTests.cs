using XlsxSharp.Excel;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Tests.Excel.Protection;

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
                ClassicAssert.IsTrue(
                    ws.Protection.AllowedElements.HasFlag(element),
                    element.ToString()
                );
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
                ClassicAssert.IsTrue(
                    ws.Protection.AllowedElements.HasFlag(element),
                    element.ToString()
                );
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
                ClassicAssert.IsTrue(
                    ws.Protection.AllowedElements.HasFlag(element),
                    element.ToString()
                );
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
                ClassicAssert.IsFalse(
                    ws.Protection.AllowedElements.HasFlag(element),
                    element.ToString()
                );
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
                ClassicAssert.IsFalse(
                    ws.Protection.AllowedElements.HasFlag(element),
                    element.ToString()
                );
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
                ClassicAssert.IsTrue(ws.Protection.IsProtected);
                ClassicAssert.AreEqual(Algorithm.SimpleHash, ws.Protection.Algorithm);

                ws.Unprotect("123");
                ws.Protect("123", Algorithm.SHA512);
                wb.Save();
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                ClassicAssert.IsTrue(ws.Protection.IsProtected);
                ClassicAssert.AreEqual(Algorithm.SHA512, ws.Protection.Algorithm);

                ClassicAssert.DoesNotThrow(() => ws.Unprotect("123"));
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
            ClassicAssert.IsTrue(p1.IsProtected);

            IXLWorksheet ws2 = ws1.CopyTo("New worksheet");
            ClassicAssert.IsFalse(ws2.Protection.IsProtected);
            XLSheetProtection p2 = ws2.Protection.CopyFrom(p1).CastTo<XLSheetProtection>();

            ClassicAssert.IsTrue(p2.IsProtected);
            ClassicAssert.IsTrue(p2.IsPasswordProtected);
            ClassicAssert.AreEqual(p1.Algorithm, p2.Algorithm);
            ClassicAssert.AreEqual(p1.PasswordHash, p2.PasswordHash);
            ClassicAssert.AreEqual(p1.Base64EncodedSalt, p2.Base64EncodedSalt);
            ClassicAssert.AreEqual(p1.SpinCount, p2.SpinCount);

            ClassicAssert.IsTrue(
                p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertColumns)
            );
            ClassicAssert.IsTrue(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertRows));
            ClassicAssert.IsFalse(
                p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertHyperlinks)
            );

            ClassicAssert.Throws<InvalidOperationException>(() => ws2.Unprotect());
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

        ClassicAssert.AreEqual(
            XLSheetProtectionElements.FormatColumns
                | XLSheetProtectionElements.FormatRows
                | XLSheetProtectionElements.SelectEverything,
            ws1.Protection.AllowedElements
        );

        ws2.Protection = ws1.Protection;

        ClassicAssert.IsFalse(ReferenceEquals(ws1.Protection, ws2.Protection));
        ClassicAssert.IsTrue(ws2.Protection.IsProtected);
        ClassicAssert.AreEqual(
            XLSheetProtectionElements.FormatColumns
                | XLSheetProtectionElements.FormatRows
                | XLSheetProtectionElements.SelectEverything,
            ws2.Protection.AllowedElements
        );
        ClassicAssert.AreEqual(
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
            ClassicAssert.IsTrue(ws.Protection.IsProtected);
            ws.Unprotect();
            ClassicAssert.IsFalse(ws.Protection.IsProtected);
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
            ClassicAssert.IsTrue(ws.Protection.IsProtected);

            // Password required
            ClassicAssert.Throws<InvalidOperationException>(() => ws.Unprotect());

            ClassicAssert.AreEqual(Algorithm.SHA512, ws.Protection.Algorithm);
            ws.Unprotect("abc");
            ClassicAssert.IsFalse(ws.Protection.IsProtected);
        }
    }
}
