using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.Misc;

public class StylesTests
{
    private static void SetupBorders(IXLRange range)
    {
        range.FirstRow().Cell(1).Style.Border.TopBorder = XLBorderStyleValues.None;
        range.FirstRow().Cell(2).Style.Border.TopBorder = XLBorderStyleValues.Thick;
        range.FirstRow().Cell(3).Style.Border.TopBorder = XLBorderStyleValues.Double;

        range.LastRow().Cell(1).Style.Border.BottomBorder = XLBorderStyleValues.None;
        range.LastRow().Cell(2).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        range.LastRow().Cell(3).Style.Border.BottomBorder = XLBorderStyleValues.Double;

        range.FirstColumn().Cell(1).Style.Border.LeftBorder = XLBorderStyleValues.None;
        range.FirstColumn().Cell(2).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
        range.FirstColumn().Cell(3).Style.Border.LeftBorder = XLBorderStyleValues.Double;

        range.LastColumn().Cell(1).Style.Border.RightBorder = XLBorderStyleValues.None;
        range.LastColumn().Cell(2).Style.Border.RightBorder = XLBorderStyleValues.Thick;
        range.LastColumn().Cell(3).Style.Border.RightBorder = XLBorderStyleValues.Double;
    }

    [Test]
    public void InsideBorderTest()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        IXLRange range = ws.Range("B2:D4");

        SetupBorders(range);

        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorderColor = XLColor.Red;

        IXLCell center = range.Cell(2, 2);

        ClassicAssert.AreEqual(XLColor.Red, center.Style.Border.TopBorderColor);
        ClassicAssert.AreEqual(XLColor.Red, center.Style.Border.BottomBorderColor);
        ClassicAssert.AreEqual(XLColor.Red, center.Style.Border.LeftBorderColor);
        ClassicAssert.AreEqual(XLColor.Red, center.Style.Border.RightBorderColor);

        ClassicAssert.AreEqual(
            XLBorderStyleValues.None,
            range.FirstRow().Cell(1).Style.Border.TopBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Thick,
            range.FirstRow().Cell(2).Style.Border.TopBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Double,
            range.FirstRow().Cell(3).Style.Border.TopBorder
        );

        ClassicAssert.AreEqual(
            XLBorderStyleValues.None,
            range.LastRow().Cell(1).Style.Border.BottomBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Thick,
            range.LastRow().Cell(2).Style.Border.BottomBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Double,
            range.LastRow().Cell(3).Style.Border.BottomBorder
        );

        ClassicAssert.AreEqual(
            XLBorderStyleValues.None,
            range.FirstColumn().Cell(1).Style.Border.LeftBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Thick,
            range.FirstColumn().Cell(2).Style.Border.LeftBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Double,
            range.FirstColumn().Cell(3).Style.Border.LeftBorder
        );

        ClassicAssert.AreEqual(
            XLBorderStyleValues.None,
            range.LastColumn().Cell(1).Style.Border.RightBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Thick,
            range.LastColumn().Cell(2).Style.Border.RightBorder
        );
        ClassicAssert.AreEqual(
            XLBorderStyleValues.Double,
            range.LastColumn().Cell(3).Style.Border.RightBorder
        );
    }

    [Test]
    public void ResolveThemeColors()
    {
        using (XLWorkbook wb = new())
        {
            string color;
            color = wb.Theme.ResolveThemeColor(XLThemeColor.Accent1).Color.ToHex();
            ClassicAssert.AreEqual("FF4F81BD", color);

            color = wb.Theme.ResolveThemeColor(XLThemeColor.Background1).Color.ToHex();
            ClassicAssert.AreEqual("FFFFFFFF", color);
        }
    }

    [Test]
    [MethodDataSource(nameof(AllThemeColors))]
    public void CanResolveAllThemeColors(XLThemeColor themeColor)
    {
        IXLTheme theme = new XLWorkbook().Theme;
        XLColor color = theme.ResolveThemeColor(themeColor);
        ClassicAssert.IsNotNull(color);
    }

    // NUnit's [Theory] auto-generated one case per enum value for an otherwise-undecorated enum
    // parameter; TUnit has no equivalent, so this replaces that data source explicitly.
    public static IEnumerable<XLThemeColor> AllThemeColors() => Enum.GetValues<XLThemeColor>();

    [Test]
    public void SetStyleViaRowReference()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Style.Font.SetFontSize(8).Font.SetFontColor(XLColor.Green).Font.SetBold(true);

            IXLRow row = ws.Row(1);
            ws.Cell(1, 1).Value = "Test";
            row.Cell(2).Value = "Test";
            row.Cells(3, 3).Value = "Test";

            foreach (IXLCell cell in ws.CellsUsed())
            {
                ClassicAssert.AreEqual(8, ws.Cell("A1").Style.Font.FontSize);
                ClassicAssert.AreEqual(XLColor.Green, ws.Cell("B1").Style.Font.FontColor);
                ClassicAssert.AreEqual(true, ws.Cell("C1").Style.Font.Bold);
            }
        }
    }
}
