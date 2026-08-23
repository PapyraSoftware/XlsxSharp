using XlsxSharp.Examples.ConditionalFormatting;

namespace XlsxSharp.Tests.Examples;

public class ConditionalFormattingTests
{
    [Test]
    public void CfColorScaleLowHigh() =>
        TestHelper.RunTestExample<CfColorScaleLowHigh>(
            @"ConditionalFormatting\CFColorScaleLowHigh.xlsx"
        );

    [Test]
    public void CfColorScaleLowMidHigh() =>
        TestHelper.RunTestExample<CfColorScaleLowMidHigh>(
            @"ConditionalFormatting\CFColorScaleLowMidHigh.xlsx"
        );

    [Test]
    public void CfColorScaleMinimumMaximum() =>
        TestHelper.RunTestExample<CfColorScaleMinimumMaximum>(
            @"ConditionalFormatting\CFColorScaleMinimumMaximum.xlsx"
        );

    [Test]
    public void CfContains() =>
        TestHelper.RunTestExample<CfContains>(@"ConditionalFormatting\CFContains.xlsx");

    [Test]
    public void CfDataBar() =>
        TestHelper.RunTestExample<CfDataBar>(@"ConditionalFormatting\CFDataBar.xlsx");

    [Test]
    public void CfDataBarNegative() =>
        TestHelper.RunTestExample<CfDataBarNegative>(
            @"ConditionalFormatting\CFDataBarNegative.xlsx"
        );

    [Test]
    public void CfEndsWith() =>
        TestHelper.RunTestExample<CfEndsWith>(@"ConditionalFormatting\CFEndsWith.xlsx");

    [Test]
    public void CfEqualsNumber() =>
        TestHelper.RunTestExample<CfEqualsNumber>(@"ConditionalFormatting\CFEqualsNumber.xlsx");

    [Test]
    public void CfEqualsString() =>
        TestHelper.RunTestExample<CfEqualsString>(@"ConditionalFormatting\CFEqualsString.xlsx");

    [Test]
    public void CfIconSet() =>
        TestHelper.RunTestExample<CfIconSet>(@"ConditionalFormatting\CFIconSet.xlsx");

    [Test]
    public void CfIsBlank() =>
        TestHelper.RunTestExample<CfIsBlank>(@"ConditionalFormatting\CFIsBlank.xlsx");

    [Test]
    public void CfIsError() =>
        TestHelper.RunTestExample<CfIsError>(@"ConditionalFormatting\CFIsError.xlsx");

    [Test]
    public void CfNotBlank() =>
        TestHelper.RunTestExample<CfNotBlank>(@"ConditionalFormatting\CFNotBlank.xlsx");

    [Test]
    public void CfNotContains() =>
        TestHelper.RunTestExample<CfNotContains>(@"ConditionalFormatting\CFNotContains.xlsx");

    [Test]
    public void CfNotEqualsNumber() =>
        TestHelper.RunTestExample<CfNotEqualsNumber>(
            @"ConditionalFormatting\CFNotEqualsNumber.xlsx"
        );

    [Test]
    public void CfNotEqualsString() =>
        TestHelper.RunTestExample<CfNotEqualsString>(
            @"ConditionalFormatting\CFNotEqualsString.xlsx"
        );

    [Test]
    public void CfNotError() =>
        TestHelper.RunTestExample<CfNotError>(@"ConditionalFormatting\CFNotError.xlsx");

    [Test]
    public void CfStartsWith() =>
        TestHelper.RunTestExample<CfStartsWith>(@"ConditionalFormatting\CFStartsWith.xlsx");

    [Test]
    public void CfMultipleConditions() =>
        TestHelper.RunTestExample<CfMultipleConditions>(
            @"ConditionalFormatting\CFMultipleConditions.xlsx"
        );

    [Test]
    public void CfStopIfTrue() =>
        TestHelper.RunTestExample<CfStopIfTrue>(@"ConditionalFormatting\CFStopIfTrue.xlsx");

    [Test]
    public void CfTop() => TestHelper.RunTestExample<CfTop>(@"ConditionalFormatting\CFTop.xlsx");

    [Test]
    public void CfBottom() =>
        TestHelper.RunTestExample<CfBottom>(@"ConditionalFormatting\CFBottom.xlsx");

    [Test]
    public void CfDatesOccurring() =>
        TestHelper.RunTestExample<CfDatesOccurring>(@"ConditionalFormatting\CFDatesOccurring.xlsx");

    [Test]
    public void CfDataBars() =>
        TestHelper.RunTestExample<CfDataBars>(@"ConditionalFormatting\CFDataBars.xlsx");
}
