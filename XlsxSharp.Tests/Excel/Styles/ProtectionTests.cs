using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

public class ProtectionTests
{
    [Test]
    [MethodDataSource(nameof(ProtectionApiSetters))]
    public void ProtectionPropertyCanBeIndividuallySet(FormatTestCase<IXLProtection> testCase)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLStyle cellFormat = ws.Cell("B2").Style;
        foreach (object testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cellFormat.Protection, testValue);
            object setValue = testCase.GetPropertyValue(cellFormat.Protection);
            ClassicAssert.AreEqual(testValue, setValue);
        }
    }

    [Test]
    public void ProtectionCanBeCopied()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLProtection targetProtection = ws.Cell("A2").Style.Protection;
        ClassicAssert.IsTrue(targetProtection.Locked);
        ClassicAssert.IsFalse(targetProtection.Hidden);
        IXLStyle source = ws.Cell("A1")
            .Style.Protection.SetLocked(false)
            .Protection.SetHidden(true);

        targetProtection = source.Protection;

        ClassicAssert.IsFalse(targetProtection.Locked);
        ClassicAssert.IsTrue(targetProtection.Hidden);
    }

    [Test]
    public void ProtectionHasEqualityComparison()
    {
        Action<IXLProtection>[] changePropertyToNonDefault =
        [
            x => x.SetLocked(false),
            x => x.SetHidden(true),
        ];

        using XLWorkbook wb = new();
        foreach (Action<IXLProtection> changeProperty in changePropertyToNonDefault)
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLProtection lhs = ws.Cell("A1").Style.Protection;
            IXLProtection rhs = ws.Cell("A2").Style.Protection;

            ClassicAssert.AreEqual(lhs, rhs);
            changeProperty(lhs);
            ClassicAssert.AreNotEqual(lhs, rhs);
        }
    }

    public static IEnumerable<FormatTestCase<IXLProtection>> ProtectionApiSetters()
    {
        bool[] boolValues = [false, true];
        yield return FormatTestCase<IXLProtection>.ForProtection(
            protection => protection.Hidden,
            (protection, value) => protection.Hidden = value,
            boolValues
        );
        yield return FormatTestCase<IXLProtection>.ForProtection(
            protection => protection.Hidden,
            (protection, value) => protection.SetHidden(value),
            boolValues
        );
        yield return FormatTestCase<IXLProtection>.ForProtection(
            protection => protection.Hidden,
            (protection, _) => protection.SetHidden(),
            true
        );

        yield return FormatTestCase<IXLProtection>.ForProtection(
            protection => protection.Locked,
            (protection, value) => protection.Locked = value,
            boolValues
        );
        yield return FormatTestCase<IXLProtection>.ForProtection(
            protection => protection.Locked,
            (protection, value) => protection.SetLocked(value),
            boolValues
        );
        yield return FormatTestCase<IXLProtection>.ForProtection(
            protection => protection.Locked,
            (protection, _) => protection.SetLocked(),
            true
        );
    }
}
