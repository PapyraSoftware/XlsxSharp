using XlsxSharp.Excel;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Tests.Excel.Protection;

public class XlWorkbookProtectionTests
{
    [Test]
    public void CanChangeProtectionAlgorithm()
    {
        using (MemoryStream ms = new())
        {
            using (Stream stream = GetProtectedWorkbookStreamWithPassword())
            using (XLWorkbook wb = new(stream))
            {
                ClassicAssert.AreEqual(Algorithm.SHA512, wb.Protection.Algorithm);
                wb.Unprotect("12345");
                wb.Protect("12345");

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                ClassicAssert.IsTrue(wb.IsPasswordProtected);
                ClassicAssert.AreEqual(Algorithm.SimpleHash, wb.Protection.Algorithm);
            }
        }
    }

    [Test]
    public void CanChangeToPasswordProtected()
    {
        using (MemoryStream ms = new())
        {
            using (Stream stream = GetProtectedWorkbookStreamWithoutPassword())
            using (XLWorkbook wb = new(stream))
            {
                wb.Unprotect();
                wb.Protection.Protect("12345");

                ClassicAssert.IsTrue(wb.Protection.IsPasswordProtected);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                ClassicAssert.IsTrue(wb.Protection.IsPasswordProtected);
                ClassicAssert.AreEqual(Algorithm.SimpleHash, wb.Protection.Algorithm);
                ClassicAssert.AreNotEqual("", wb.Protection.PasswordHash);
            }
        }
    }

    [Test]
    public void CanChangeToProtectedWithoutPassword()
    {
        using (MemoryStream ms = new())
        {
            using (Stream stream = GetProtectedWorkbookStreamWithPassword())
            using (XLWorkbook wb = new(stream))
            {
                wb.Unprotect("12345");
                wb.Protection.Protect();

                ClassicAssert.IsFalse(wb.Protection.IsPasswordProtected);
                ClassicAssert.IsTrue(wb.Protection.IsProtected);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                ClassicAssert.IsFalse(wb.Protection.IsPasswordProtected);
                ClassicAssert.IsTrue(wb.Protection.IsProtected);
                ClassicAssert.AreEqual(Algorithm.SimpleHash, wb.Protection.Algorithm);
                ClassicAssert.AreEqual("", wb.Protection.PasswordHash);
            }
        }
    }

    [Test]
    public void CannotUnprotectIfNoPassword()
    {
        using (Stream stream = GetProtectedWorkbookStreamWithoutPassword())
        using (XLWorkbook wb = new(stream))
        {
            ArgumentException? ex = ClassicAssert.Throws<ArgumentException>(() =>
                wb.Unprotect("dummy password")
            );
            ClassicAssert.AreEqual("Invalid password", ex.Message);
        }
    }

    [Test]
    public void CannotUnprotectWithoutPassword()
    {
        using (Stream stream = GetProtectedWorkbookStreamWithPassword())
        using (XLWorkbook wb = new(stream))
        {
            InvalidOperationException? ex = ClassicAssert.Throws<InvalidOperationException>(() =>
                wb.Unprotect()
            );
            ClassicAssert.AreEqual("The workbook structure is password protected", ex.Message);
        }
    }

    [Test]
    [MethodDataSource(nameof(AllAlgorithms))]
    public void CanProtectWithPassword(Algorithm algorithm)
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                wb.AddWorksheet();

                ClassicAssert.IsFalse(wb.Protection.IsProtected);

                wb.Protection.Protect("12345", algorithm);

                wb.Protection.AllowNone();
                ClassicAssert.IsFalse(
                    wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)
                );
                ClassicAssert.IsFalse(
                    wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)
                );

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                ClassicAssert.IsTrue(wb.Protection.IsPasswordProtected);
                ClassicAssert.IsTrue(wb.Protection.IsProtected);

                ClassicAssert.AreEqual(algorithm, wb.Protection.Algorithm);
                ClassicAssert.AreNotEqual("", wb.Protection.PasswordHash);

                ClassicAssert.IsFalse(
                    wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)
                );
                ClassicAssert.IsFalse(
                    wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)
                );

                ArgumentException? ex = ClassicAssert.Throws<ArgumentException>(() =>
                    wb.Unprotect("dummy password")
                );
                ClassicAssert.AreEqual("Invalid password", ex.Message);

                wb.Protection.Unprotect("12345");

                wb.Save();
            }
        }
    }

    [Test]
    public void CanUnprotectWithoutPassword()
    {
        using (MemoryStream ms = new())
        {
            using (Stream stream = GetProtectedWorkbookStreamWithoutPassword())
            using (XLWorkbook wb = new(stream))
            {
                // Unprotect without password
                wb.Unprotect();

                ClassicAssert.IsFalse(wb.Protection.IsProtected);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                ClassicAssert.IsFalse(wb.Protection.IsProtected);
            }
        }
    }

    [Test]
    public void CanUnprotectWithPassword()
    {
        using (MemoryStream ms = new())
        {
            using (Stream stream = GetProtectedWorkbookStreamWithPassword())
            using (XLWorkbook wb = new(stream))
            {
                // Unprotect with password
                wb.Unprotect("12345");

                ClassicAssert.IsFalse(wb.Protection.IsProtected);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                ClassicAssert.IsFalse(wb.Protection.IsProtected);
            }
        }
    }

    [Test]
    public void CopyProtectionFromAnotherWorkbook()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Misc\WorkbookProtection.xlsx")
            )
        )
        using (XLWorkbook wb1 = new(stream))
        using (XLWorkbook wb2 = new())
        {
            wb2.AddWorksheet();

            XLWorkbookProtection p1 = wb1.Protection.CastTo<XLWorkbookProtection>();
            ClassicAssert.IsTrue(p1.IsProtected);

            ClassicAssert.IsFalse(wb2.Protection.IsProtected);
            XLWorkbookProtection p2 = wb2
                .Protection.CopyFrom(wb1.Protection)
                .CastTo<XLWorkbookProtection>();

            ClassicAssert.IsTrue(p2.IsProtected);
            ClassicAssert.IsTrue(p2.IsPasswordProtected);
            ClassicAssert.AreEqual(p1.Algorithm, p2.Algorithm);
            ClassicAssert.AreEqual(p1.PasswordHash, p2.PasswordHash);
            ClassicAssert.AreEqual(p1.Base64EncodedSalt, p2.Base64EncodedSalt);
            ClassicAssert.AreEqual(p1.SpinCount, p2.SpinCount);

            ClassicAssert.IsTrue(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows));
            ClassicAssert.IsFalse(
                p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)
            );

            ClassicAssert.Throws<InvalidOperationException>(() => wb2.Unprotect());
            wb2.Unprotect("Abc@123");
        }
    }

    [Test]
    public void IxlProtectableTests()
    {
        using XLWorkbook wb = new();
        Enumerable.Range(1, 5).ForEach(i => wb.AddWorksheet());

        List<IXLProtectable> list = [wb, .. wb.Worksheets];

        list.ForEach(el => el.Protect());

        list.ForEach(el => ClassicAssert.IsTrue(el.IsProtected));
        list.ForEach(el => ClassicAssert.IsFalse(el.IsPasswordProtected));

        list.ForEach(el => el.Unprotect());

        list.ForEach(el => ClassicAssert.IsFalse(el.IsProtected));
        list.ForEach(el => ClassicAssert.IsFalse(el.IsPasswordProtected));

        list.ForEach(el => el.Protect("password"));

        list.ForEach(el => ClassicAssert.IsTrue(el.IsProtected));
        list.ForEach(el => ClassicAssert.IsTrue(el.IsPasswordProtected));

        list.ForEach(el => el.Unprotect("password"));

        list.ForEach(el => ClassicAssert.IsFalse(el.IsProtected));
        list.ForEach(el => ClassicAssert.IsFalse(el.IsPasswordProtected));
    }

    [Test]
    public void LoadProtectionWithoutPasswordFromFile()
    {
        using (Stream stream = GetProtectedWorkbookStreamWithoutPassword())
        using (XLWorkbook wb = new(stream))
        {
            ClassicAssert.IsFalse(wb.Protection.IsPasswordProtected);
            ClassicAssert.IsTrue(wb.Protection.IsProtected);
            ClassicAssert.AreEqual("", wb.Protection.PasswordHash);
            ClassicAssert.IsTrue(
                wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)
            );
            ClassicAssert.IsFalse(
                wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)
            );
        }
    }

    [Test]
    public void LoadProtectionWithPasswordFromFile()
    {
        using (Stream stream = GetProtectedWorkbookStreamWithPassword())
        using (XLWorkbook wb = new(stream))
        {
            ClassicAssert.IsTrue(wb.Protection.IsPasswordProtected);
            ClassicAssert.AreNotEqual("", wb.Protection.PasswordHash);
            ClassicAssert.IsTrue(
                wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)
            );
            ClassicAssert.IsFalse(
                wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)
            );
        }
    }

    [Test]
    public void SetWorkbookProtectionCloning()
    {
        XLWorkbook wb1 = new();
        XLWorkbook wb2 = new();

        wb1.AddWorksheet();
        wb2.AddWorksheet();

        wb1.Protect("123", Algorithm.SHA512)
            .AllowElement(XLWorkbookProtectionElements.Windows)
            .DisallowElement(XLWorkbookProtectionElements.Structure);

        ClassicAssert.IsTrue(wb1.Protection.IsProtected);

        ClassicAssert.AreEqual(
            XLWorkbookProtectionElements.Windows,
            wb1.Protection.AllowedElements
        );

        wb2.Protection = wb1.Protection;

        ClassicAssert.IsFalse(ReferenceEquals(wb1.Protection, wb2.Protection));
        ClassicAssert.IsTrue(wb2.Protection.IsProtected);
        ClassicAssert.AreEqual(
            XLWorkbookProtectionElements.Windows,
            wb2.Protection.AllowedElements
        );
        ClassicAssert.AreEqual(wb1.Protection.PasswordHash, wb2.Protection.PasswordHash);
    }

    private static Stream GetProtectedWorkbookStreamWithoutPassword() =>
        TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithoutpassword.xlsx")
        );

    private static Stream GetProtectedWorkbookStreamWithPassword() =>
        TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithpassword.xlsx")
        );

    // NUnit's [Theory] auto-generated one case per enum value for an otherwise-undecorated enum
    // parameter; TUnit has no equivalent, so this replaces that data source explicitly.
    internal static IEnumerable<Algorithm> AllAlgorithms() => Enum.GetValues<Algorithm>();
}
