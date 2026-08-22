using NUnit.Framework;
using XlsxSharp.Examples.ConditionalFormatting;

namespace XlsxSharp.Tests.Examples;

[TestFixture]
public class ConditionalFormattingTests
{
    [Test]
    public void CFColorScaleLowHigh() =>
        TestHelper.RunTestExample<CfColorScaleLowHigh>(
            @"ConditionalFormatting\CFColorScaleLowHigh.xlsx"
        );

    [Test]
    public void CFColorScaleLowMidHigh() =>
        TestHelper.RunTestExample<CfColorScaleLowMidHigh>(
            @"ConditionalFormatting\CFColorScaleLowMidHigh.xlsx"
        );

    [Test]
    public void CFColorScaleMinimumMaximum() =>
        TestHelper.RunTestExample<CfColorScaleMinimumMaximum>(
            @"ConditionalFormatting\CFColorScaleMinimumMaximum.xlsx"
        );

    [Test]
    public void CFContains() =>
        TestHelper.RunTestExample<CfContains>(@"ConditionalFormatting\CFContains.xlsx");

    [Test]
    public void CFDataBar() =>
        TestHelper.RunTestExample<CfDataBar>(@"ConditionalFormatting\CFDataBar.xlsx");

    [Test]
    public void CFDataBarNegative() =>
        TestHelper.RunTestExample<CfDataBarNegative>(
            @"ConditionalFormatting\CFDataBarNegative.xlsx"
        );

    [Test]
    public void CFEndsWith() =>
        TestHelper.RunTestExample<CfEndsWith>(@"ConditionalFormatting\CFEndsWith.xlsx");

    [Test]
    public void CFEqualsNumber() =>
        TestHelper.RunTestExample<CfEqualsNumber>(@"ConditionalFormatting\CFEqualsNumber.xlsx");

    [Test]
    public void CFEqualsString() =>
        TestHelper.RunTestExample<CfEqualsString>(@"ConditionalFormatting\CFEqualsString.xlsx");

    [Test]
    public void CFIconSet() =>
        TestHelper.RunTestExample<CfIconSet>(@"ConditionalFormatting\CFIconSet.xlsx");

    [Test]
    public void CFIsBlank() =>
        TestHelper.RunTestExample<CfIsBlank>(@"ConditionalFormatting\CFIsBlank.xlsx");

    [Test]
    public void CFIsError() =>
        TestHelper.RunTestExample<CfIsError>(@"ConditionalFormatting\CFIsError.xlsx");

    [Test]
    public void CFNotBlank() =>
        TestHelper.RunTestExample<CfNotBlank>(@"ConditionalFormatting\CFNotBlank.xlsx");

    [Test]
    public void CFNotContains() =>
        TestHelper.RunTestExample<CfNotContains>(@"ConditionalFormatting\CFNotContains.xlsx");

    [Test]
    public void CFNotEqualsNumber() =>
        TestHelper.RunTestExample<CfNotEqualsNumber>(
            @"ConditionalFormatting\CFNotEqualsNumber.xlsx"
        );

    [Test]
    public void CFNotEqualsString() =>
        TestHelper.RunTestExample<CfNotEqualsString>(
            @"ConditionalFormatting\CFNotEqualsString.xlsx"
        );

    [Test]
    public void CFNotError() =>
        TestHelper.RunTestExample<CfNotError>(@"ConditionalFormatting\CFNotError.xlsx");

    [Test]
    public void CFStartsWith() =>
        TestHelper.RunTestExample<CfStartsWith>(@"ConditionalFormatting\CFStartsWith.xlsx");

    [Test]
    public void CFMultipleConditions() =>
        TestHelper.RunTestExample<CfMultipleConditions>(
            @"ConditionalFormatting\CFMultipleConditions.xlsx"
        );

    [Test]
    public void CFStopIfTrue() =>
        TestHelper.RunTestExample<CfStopIfTrue>(@"ConditionalFormatting\CFStopIfTrue.xlsx");

    [Test]
    public void CFTop() => TestHelper.RunTestExample<CfTop>(@"ConditionalFormatting\CFTop.xlsx");

    [Test]
    public void CFBottom() =>
        TestHelper.RunTestExample<CfBottom>(@"ConditionalFormatting\CFBottom.xlsx");

    [Test]
    public void CFDatesOccurring() =>
        TestHelper.RunTestExample<CfDatesOccurring>(@"ConditionalFormatting\CFDatesOccurring.xlsx");

    [Test]
    public void CFDataBars() =>
        TestHelper.RunTestExample<CfDataBars>(@"ConditionalFormatting\CFDataBars.xlsx");
}
