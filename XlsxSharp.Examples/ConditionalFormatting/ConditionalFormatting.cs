using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.ConditionalFormats;

namespace XlsxSharp.Examples.ConditionalFormatting;

public class CfColorScaleLowMidHigh : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.Red)
            .Midpoint(XLCFContentType.Percent, "50", XLColor.Yellow)
            .HighestValue(XLColor.Green);

        workbook.SaveAs(filePath);
    }
}

public class CfColorScaleLowHigh : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .ColorScale()
            .Minimum(XLCFContentType.Number, "2", XLColor.Red)
            .Maximum(XLCFContentType.Percentile, "90", XLColor.Green);

        workbook.SaveAs(filePath);
    }
}

public class CfColorScaleMinimumMaximum : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.FromHtml("#FFFF7128"))
            .HighestValue(XLColor.FromHtml("#FFFFEF9C"));

        workbook.SaveAs(filePath);
    }
}

public class CfStartsWith : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue("Hellos")
            .CellBelow()
            .SetValue("Hell")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenStartsWith("Hell")
            .Fill.SetBackgroundColor(XLColor.Red)
            .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XLColor.Blue)
            .Font.SetBold();

        workbook.SaveAs(filePath);
    }
}

public class CfEndsWith : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue("Hellos")
            .CellBelow()
            .SetValue("Hell")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenEndsWith("ll")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfIsBlank : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue(Blank.Value)
            .CellBelow()
            .SetValue("")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed().AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfNotBlank : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue(Blank.Value)
            .CellBelow()
            .SetValue("")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed().AddConditionalFormat().WhenNotBlank().Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfIsError : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetFormulaA1("1/0")
            .CellBelow()
            .SetFormulaA1("1/0")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed().AddConditionalFormat().WhenIsError().Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfNotError : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetFormulaA1("1/0")
            .CellBelow()
            .SetFormulaA1("1/0")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed().AddConditionalFormat().WhenNotError().Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfContains : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue("Hellos")
            .CellBelow()
            .SetValue("Hell")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenContains("Hell")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfNotContains : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue("Hellos")
            .CellBelow()
            .SetValue("Hell")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenNotContains("Hell")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfEqualsString : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue("Hellos")
            .CellBelow()
            .SetValue("Hell")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenEquals("Hell")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfEqualsNumber : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed().AddConditionalFormat().WhenEquals(2).Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfNotEqualsString : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue("Hello")
            .CellBelow()
            .SetValue("Hellos")
            .CellBelow()
            .SetValue("Hell")
            .CellBelow()
            .SetValue("Holl");

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenNotEquals("Hell")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfNotEqualsNumber : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed().AddConditionalFormat().WhenNotEquals(2).Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfGreaterThan : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenGreaterThan("2")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfEqualOrGreaterThan : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenEqualOrGreaterThan("2")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfLessThan : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenLessThan("2")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfEqualOrLessThan : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenEqualOrLessThan("2")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfBetween : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenBetween("2", "3")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfNotBetween : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenNotBetween("2", "3")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfUnique : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed().AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfDuplicate : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenIsDuplicate()
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfIsTrue : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenIsTrue("TRUE")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfTop : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed().AddConditionalFormat().WhenIsTop(2).Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfBottom : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenIsBottom(10, XLTopBottomType.Percent)
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfDataBar : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .DataBar(XLColor.Red, true)
            .LowestValue()
            .Maximum(XLCFContentType.Percent, "100");

        workbook.SaveAs(filePath);
    }
}

public class CfDataBarNegative : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.Cell(1, 1)
            .SetValue(-1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.Range(ws.Cell(1, 1), ws.Cell(4, 1))
            .AddConditionalFormat()
            .DataBar(XLColor.Green, XLColor.Red, showBarOnly: false)
            .LowestValue()
            .HighestValue();

        ws.Cell(1, 3)
            .SetValue(-20)
            .CellBelow()
            .SetValue(40)
            .CellBelow()
            .SetValue(-60)
            .CellBelow()
            .SetValue(30);

        ws.Range(ws.Cell(1, 3), ws.Cell(4, 3))
            .AddConditionalFormat()
            .DataBar(XLColor.Green, XLColor.Red, showBarOnly: true)
            .Minimum(XLCFContentType.Number, -100)
            .Maximum(XLCFContentType.Number, 100);

        workbook.SaveAs(filePath);
    }
}

public class CfIconSet : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .IconSet(XLIconSetStyle.ThreeTrafficLights2, true, true)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "2", XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "3", XLCFContentType.Number);

        workbook.SaveAs(filePath);
    }
}

public class CfTwoConditions : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed()
            .AddConditionalFormat()
            .IconSet(XLIconSetStyle.ThreeTrafficLights2, true, true)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "2", XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "3", XLCFContentType.Number);

        ws.RangeUsed()
            .AddConditionalFormat()
            .WhenContains("1")
            .Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfInsertRows : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.Cell(2, 1)
            .SetValue(1)
            .CellRight()
            .SetValue(1)
            .CellRight()
            .SetValue(2)
            .CellRight()
            .SetValue(3);

        IXLRange? range = ws.RangeUsed();
        range.AddConditionalFormat().WhenEquals("1").Font.SetBold();
        range.InsertRowsAbove(1);

        workbook.SaveAs(filePath);
    }
}

public class CfTest : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(1)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(4);

        ws.RangeUsed()
            .AddConditionalFormat()
            .DataBar(XLColor.Red, XLColor.Green)
            .LowestValue()
            .HighestValue();

        workbook.SaveAs(filePath);
    }
}

public class CfMultipleConditions : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        IXLRange range = ws.Range("A1:A10");
        range.AddConditionalFormat().WhenEquals("3").Fill.SetBackgroundColor(XLColor.Blue);
        range.AddConditionalFormat().WhenEquals("2").Fill.SetBackgroundColor(XLColor.Green);
        range.AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Red);

        workbook.SaveAs(filePath);
    }
}

public class CfStopIfTrue : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

        ws.FirstCell()
            .SetValue(6)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3);

        ws.RangeUsed().AddConditionalFormat().SetStopIfTrue().WhenGreaterThan(5);

        ws.RangeUsed()
            .AddConditionalFormat()
            .IconSet(XLIconSetStyle.ThreeTrafficLights2, true, true)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "2", XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "3", XLCFContentType.Number);

        workbook.SaveAs(filePath);
    }
}

public class CfDatesOccurring : IXLExample
{
    public void Create(string filePath)
    {
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet ws = workbook.AddWorksheet("Sheet1");

            IXLRange range = ws.Range("A1:A10");
            range
                .AddConditionalFormat()
                .WhenDateIs(XLTimePeriod.Tomorrow)
                .Fill.SetBackgroundColor(XLColor.GrannySmithApple);

            range
                .AddConditionalFormat()
                .WhenDateIs(XLTimePeriod.Yesterday)
                .Fill.SetBackgroundColor(XLColor.Orange);

            range
                .AddConditionalFormat()
                .WhenDateIs(XLTimePeriod.InTheLast7Days)
                .Fill.SetBackgroundColor(XLColor.Blue);

            range
                .AddConditionalFormat()
                .WhenDateIs(XLTimePeriod.ThisMonth)
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }
}

public class CfDataBars : IXLExample
{
    public void Create(string filePath)
    {
        using XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.AddWorksheet();

        ws.Range("A2:F3").Value = 1;
        ws.Range("A4:F4").Value = 2;
        ws.Range("A5:F5").Value = 3;
        ws.Range("A6:F6").Value = 4;

        ws.Cell("A1").Value = "Automatic";
        ws.Range("A2:A6").AddConditionalFormat().DataBar(XLColor.Amber);

        ws.Cell("B1").Value = "Lowest/Highest";
        ws.Range("B2:B6")
            .AddConditionalFormat()
            .DataBar(XLColor.BallBlue)
            .LowestValue()
            .HighestValue();

        ws.Cell("C1").Value = "Value";
        ws.Range("C2:C6")
            .AddConditionalFormat()
            .DataBar(XLColor.Cadet)
            .Minimum(XLCFContentType.Number, 0)
            .Maximum(XLCFContentType.Number, 10);

        ws.Cell("D1").Value = "Percent";
        ws.Range("D2:D6")
            .AddConditionalFormat()
            .DataBar(XLColor.Desert)
            .Minimum(XLCFContentType.Percent, 50)
            .Maximum(XLCFContentType.Percent, 100);

        ws.Cell("E1").Value = "Formula";
        ws.Range("E2:E6")
            .AddConditionalFormat()
            .DataBar(XLColor.Ecru)
            .Minimum(XLCFContentType.Formula, "-SUM($A$2:$E$2)")
            .Maximum(XLCFContentType.Formula, "SUM($A$6:$E$6)");

        ws.Cell("F1").Value = "Percentile";
        ws.Range("F2:F6")
            .AddConditionalFormat()
            .DataBar(XLColor.Fandango)
            .Minimum(XLCFContentType.Percentile, 30)
            .Maximum(XLCFContentType.Percentile, 70);

        workbook.SaveAs(filePath);
    }
}
