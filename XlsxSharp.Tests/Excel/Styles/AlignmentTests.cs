using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.Styles;

[TestFixture]
public class AlignmentTests
{
    [Test]
    public void TextRotationCanBeFromMinus90To90DegreesAnd255ForVerticalLayout() =>
        TestHelper.CreateAndCompare(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();
                ws.ColumnWidth = 10;
                ws.Cell(1, 1).SetValue("Vertical: 255").Style.Alignment.SetTextRotation(255);

                for (int angle = -90; angle <= +90; angle += 10)
                {
                    int column = (angle + 90) / 10 + 2;
                    IXLCell cell = ws.Cell(1, column);
                    cell.Value = $"Rotation: {angle}";
                    cell.Style.Alignment.TextRotation = angle;
                }
            },
            @"Other\Styles\Alignment\TextRotation.xlsx"
        );

    [Test]
    public void TextRotationIsConvertedOnLoadToMinus90To90Degrees() =>
        TestHelper.LoadAndAssert(
            wb =>
            {
                IXLWorksheet ws = wb.Worksheets.Single();
                Assert.AreEqual(255, ws.Cell(1, 1).Style.Alignment.TextRotation);
                for (int column = 2; column < 21; ++column)
                {
                    int expectedAngle = (column - 2) * 10 - 90;
                    Assert.AreEqual(expectedAngle, ws.Cell(1, column).Style.Alignment.TextRotation);
                }
            },
            @"Other\Styles\Alignment\TextRotation.xlsx"
        );

    [TestCase(91)]
    [TestCase(-91)]
    [TestCase(254)]
    [TestCase(256)]
    public void TextRotationOutsideBoundsThrowsException(int textRotation) =>
        Assert.Throws<ArgumentOutOfRangeException>(() =>
        {
            using XLWorkbook wb = new();
            IXLWorksheet ws = wb.AddWorksheet();
            ws.FirstCell().Style.Alignment.TextRotation = textRotation;
        });

    [Test]
    [TestCaseSource(nameof(AlignmentApiSetters))]
    public void AlignmentPropertyCanBeIndividuallySet(FormatTestCase<IXLAlignment> testCase)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLStyle cellFormat = ws.Cell("B2").Style;
        foreach (object testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cellFormat.Alignment, testValue);
            object setValue = testCase.GetPropertyValue(cellFormat.Alignment);
            Assert.AreEqual(testValue, setValue);
        }
    }

    [Test]
    [TestCaseSource(nameof(AlignmentApiSettersLimits))]
    public void AlignmentPropertyLimitsThrowException<T>(
        Action<IXLAlignment, T> setter,
        params T[] invalidValues
    )
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLAlignment alignment = ws.Cell("A1").Style.Alignment;

        foreach (T invalidValue in invalidValues)
        {
            Assert.That(
                () => setter(alignment, invalidValue),
                Throws.TypeOf<ArgumentOutOfRangeException>()
            );
        }
    }

    [Test]
    public void TextRotationIsConnectedWithTopToBottom()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLAlignment alignment = ws.Cell("A1").Style.Alignment;

        alignment.TopToBottom = true;
        Assert.AreEqual(255, alignment.TextRotation);

        alignment.TopToBottom = false;
        Assert.AreEqual(0, alignment.TextRotation);

        alignment.TextRotation = 0;
        Assert.IsFalse(alignment.TopToBottom);

        alignment.TextRotation = 10;
        Assert.IsFalse(alignment.TopToBottom);

        alignment.TextRotation = 255;
        Assert.IsTrue(alignment.TopToBottom);
    }

    [Test]
    public void AlignmentCanBeCopied()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle source = ws.Cell("A1")
            .Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)
            .Alignment.SetVertical(XLAlignmentVerticalValues.Top)
            .Alignment.SetIndent(2)
            .Alignment.SetJustifyLastLine()
            .Alignment.SetReadingOrder(XLAlignmentReadingOrderValues.LeftToRight)
            .Alignment.SetRelativeIndent(3)
            .Alignment.SetShrinkToFit()
            .Alignment.SetTextRotation(12)
            .Alignment.SetWrapText();

        ws.Cell("A2").Style.Alignment = source.Alignment;

        IXLAlignment targetAlignment = ws.Cell("A2").Style.Alignment;
        Assert.AreEqual(XLAlignmentHorizontalValues.Right, targetAlignment.Horizontal);
        Assert.AreEqual(XLAlignmentVerticalValues.Top, targetAlignment.Vertical);
        Assert.AreEqual(2, targetAlignment.Indent);
        Assert.IsTrue(targetAlignment.JustifyLastLine);
        Assert.AreEqual(XLAlignmentReadingOrderValues.LeftToRight, targetAlignment.ReadingOrder);
        Assert.AreEqual(3, targetAlignment.RelativeIndent);
        Assert.IsTrue(targetAlignment.ShrinkToFit);
        Assert.AreEqual(12, targetAlignment.TextRotation);
        Assert.IsTrue(targetAlignment.WrapText);
    }

    [Test]
    public void AlignmentHasEqualityComparison()
    {
        Action<IXLAlignment>[] changePropertyToNonDefault =
        [
            x => x.SetHorizontal(XLAlignmentHorizontalValues.Right),
            x => x.SetVertical(XLAlignmentVerticalValues.Top),
            x => x.SetIndent(2),
            x => x.SetJustifyLastLine(),
            x => x.SetReadingOrder(XLAlignmentReadingOrderValues.LeftToRight),
            x => x.SetRelativeIndent(3),
            x => x.SetShrinkToFit(),
            x => x.SetTextRotation(12),
            x => x.SetWrapText(),
        ];

        using XLWorkbook wb = new();
        foreach (Action<IXLAlignment> changeProperty in changePropertyToNonDefault)
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLAlignment source = ws.Cell("A1").Style.Alignment;
            IXLAlignment target = ws.Cell("A2").Style.Alignment;

            Assert.AreEqual(source, target);
            changeProperty(source);
            Assert.AreNotEqual(source, target);
        }
    }

    private static IEnumerable<FormatTestCase<IXLAlignment>> AlignmentApiSetters()
    {
        XLAlignmentHorizontalValues[] hAlignValues =
            EnumPolyfill.GetValues<XLAlignmentHorizontalValues>();
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.Horizontal,
            (align, value) => align.Horizontal = value,
            hAlignValues
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.Horizontal,
            (align, value) => align.SetHorizontal(value),
            hAlignValues
        );

        XLAlignmentVerticalValues[] vAlignValues =
            EnumPolyfill.GetValues<XLAlignmentVerticalValues>();
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.Vertical,
            (align, value) => align.Vertical = value,
            vAlignValues
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.Vertical,
            (align, value) => align.SetVertical(value),
            vAlignValues
        );

        int[] testIndentValues = [0, 1, 5, 255];
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.Indent,
            (align, value) => align.Indent = value,
            testIndentValues
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.Indent,
            (align, value) => align.SetIndent(value),
            testIndentValues
        );

        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.JustifyLastLine,
            (align, value) => align.JustifyLastLine = value,
            false,
            true
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.JustifyLastLine,
            (align, value) => align.SetJustifyLastLine(value),
            false,
            true
        );

        XLAlignmentReadingOrderValues[] readingOrderValues =
            EnumPolyfill.GetValues<XLAlignmentReadingOrderValues>();
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.ReadingOrder,
            (align, value) => align.ReadingOrder = value,
            readingOrderValues
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.ReadingOrder,
            (align, value) => align.SetReadingOrder(value),
            readingOrderValues
        );

        // This is likely a property slated for deletion, but at least test for now
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.RelativeIndent,
            (align, value) => align.RelativeIndent = value,
            -10,
            0,
            10
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.RelativeIndent,
            (align, value) => align.SetRelativeIndent(value),
            -10,
            0,
            10
        );

        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.ShrinkToFit,
            (align, value) => align.ShrinkToFit = value,
            false,
            true
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.ShrinkToFit,
            (align, value) => align.SetShrinkToFit(value),
            false,
            true
        );

        int[] testTextRotationValues = [-90, -5, 0, 5, 90, 255];
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.TextRotation,
            (align, value) => align.TextRotation = value,
            testTextRotationValues
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.TextRotation,
            (align, value) => align.SetTextRotation(value),
            testTextRotationValues
        );

        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.WrapText,
            (align, value) => align.WrapText = value,
            false,
            true
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.WrapText,
            (align, value) => align.SetWrapText(value),
            false,
            true
        );

        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.TopToBottom,
            (align, value) => align.TopToBottom = value,
            false,
            true
        );
        yield return FormatTestCase<IXLAlignment>.ForAlignment(
            align => align.TopToBottom,
            (align, value) => align.SetTopToBottom(value),
            false,
            true
        );
    }

    private static IEnumerable<TestCaseData> AlignmentApiSettersLimits()
    {
        yield return MakeCase((align, value) => align.Indent = value, -1, 256);
        yield return MakeCase((align, value) => align.TextRotation = value, -91, 91, 254, 256);
        yield break;

        static TestCaseData MakeCase<T>(Action<IXLAlignment, T> setter, params T[] invalidValues)
        {
            return new TestCaseData(setter, invalidValues);
        }
    }
}
