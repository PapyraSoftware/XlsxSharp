using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Styles;

public class StyleFill : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Style Fill");

        int co = 2;
        int ro = 1;

        ws.Cell(++ro, co + 1).Value = "BackgroundColor = Red";
        ws.Cell(ro, co).Style.Fill.BackgroundColor = XLColor.Red;

        ws.Cell(++ro, co + 1).Value =
            "PatternType = DarkTrellis; PatternColor = Orange; BackgroundColor = Blue";
        ws.Cell(ro, co).Style.Fill.PatternType = XLFillPatternValues.DarkTrellis;
        ws.Cell(ro, co).Style.Fill.PatternColor = XLColor.Orange;
        ws.Cell(ro, co).Style.Fill.BackgroundColor = XLColor.Blue;

        workbook.SaveAs(filePath);
    }
}
