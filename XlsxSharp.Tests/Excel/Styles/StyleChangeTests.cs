using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

public class StyleChangeTests
{
    [Test]
    public void ChangeFontColorDoesNotAffectOtherProperties()
    {
        using (XLWorkbook wb = new())
        {
            // Arrange
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell a1 = ws.Cell("A1");
            IXLCell a2 = ws.Cell("A2");
            IXLCell b1 = ws.Cell("B1");
            IXLCell b2 = ws.Cell("B2");

            ws.Range("A1:B2").Value = "Test";

            a1.Style.Fill.BackgroundColor = XLColor.Red;
            a2.Style.Fill.BackgroundColor = XLColor.Green;
            b1.Style.Fill.BackgroundColor = XLColor.Blue;
            b2.Style.Fill.BackgroundColor = XLColor.Pink;

            a1.Style.Font.FontName = "Arial";
            a2.Style.Font.FontName = "Times New Roman";
            b1.Style.Font.FontName = "Calibri";
            b2.Style.Font.FontName = "Cambria";

            // Act
            ws.Range("A1:B2").Style.Font.FontColor = XLColor.PowderBlue;

            //Assert
            ClassicAssert.AreEqual(XLColor.Red, ws.Cell("A1").Style.Fill.BackgroundColor);
            ClassicAssert.AreEqual(XLColor.Green, ws.Cell("A2").Style.Fill.BackgroundColor);
            ClassicAssert.AreEqual(XLColor.Blue, ws.Cell("B1").Style.Fill.BackgroundColor);
            ClassicAssert.AreEqual(XLColor.Pink, ws.Cell("B2").Style.Fill.BackgroundColor);

            ClassicAssert.AreEqual("Arial", ws.Cell("A1").Style.Font.FontName);
            ClassicAssert.AreEqual("Times New Roman", ws.Cell("A2").Style.Font.FontName);
            ClassicAssert.AreEqual("Calibri", ws.Cell("B1").Style.Font.FontName);
            ClassicAssert.AreEqual("Cambria", ws.Cell("B2").Style.Font.FontName);

            ClassicAssert.AreEqual(XLColor.PowderBlue, ws.Cell("A1").Style.Font.FontColor);
            ClassicAssert.AreEqual(XLColor.PowderBlue, ws.Cell("A2").Style.Font.FontColor);
            ClassicAssert.AreEqual(XLColor.PowderBlue, ws.Cell("B1").Style.Font.FontColor);
            ClassicAssert.AreEqual(XLColor.PowderBlue, ws.Cell("B2").Style.Font.FontColor);
        }
    }

    [Test]
    public void ChangeStyleAlignment()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle style = ws.Style;

        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

        ClassicAssert.AreEqual(XLAlignmentHorizontalValues.Justify, style.Alignment.Horizontal);
    }

    [Test]
    public void ChangeStyleBorder()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle style = ws.Style;

        style.Border.DiagonalBorder = XLBorderStyleValues.Double;

        ClassicAssert.AreEqual(XLBorderStyleValues.Double, style.Border.DiagonalBorder);
    }

    [Test]
    public void ChangeStyleFill()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle style = ws.Style;

        style.Fill.BackgroundColor = XLColor.Red;

        ClassicAssert.AreEqual(XLColor.Red, style.Fill.BackgroundColor);
    }

    [Test]
    public void ChangeStyleFont()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle style = ws.Style;

        style.Font.FontSize = 50;

        ClassicAssert.AreEqual(50, style.Font.FontSize);
    }

    [Test]
    public void ChangeStyleNumberFormat()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle style = ws.Style;

        style.NumberFormat.Format = "YYYY";

        ClassicAssert.AreEqual("YYYY", style.NumberFormat.Format);
    }

    [Test]
    public void ChangeStyleProtection()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle style = ws.Style;

        style.Protection.Hidden = true;

        ClassicAssert.AreEqual(true, style.Protection.Hidden);
    }

    [Test]
    public void ChangeAttachedStyleAlignment()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell a1 = ws.Cell("A1");

            a1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

            ClassicAssert.AreEqual(
                XLAlignmentHorizontalValues.Justify,
                a1.Style.Alignment.Horizontal
            );
        }
    }
}
