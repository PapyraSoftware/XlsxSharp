using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Styles;

public class StyleAlignment : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Style Alignment");

        int co = 2;
        int ro = 1;

        ws.Cell(++ro, co).Value = "Horizontal = Right";
        ws.Cell(ro, co).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

        ws.Cell(++ro, co).Value = "Indent = 2";
        ws.Cell(ro, co).Style.Alignment.Indent = 2;

        ws.Cell(++ro, co).Value = "JustifyLastLine = true";
        ws.Cell(ro, co).Style.Alignment.JustifyLastLine = true;

        ws.Cell(++ro, co).Value = "ReadingOrder = ContextDependent";
        ws.Cell(ro, co).Style.Alignment.ReadingOrder =
            XLAlignmentReadingOrderValues.ContextDependent;

        ws.Cell(++ro, co).Value = "RelativeIndent = 2";
        ws.Cell(ro, co).Style.Alignment.RelativeIndent = 2;

        ws.Cell(++ro, co).Value = "ShrinkToFit = true";
        ws.Cell(ro, co).Style.Alignment.ShrinkToFit = true;

        ws.Cell(++ro, co).Value = "TextRotation = 45";
        ws.Cell(ro, co).Style.Alignment.TextRotation = 45;

        ws.Cell(++ro, co).Value = "TopToBottom = true";
        ws.Cell(ro, co).Style.Alignment.TopToBottom = true;

        ws.Cell(++ro, co).Value = "Vertical = Center";
        ws.Cell(ro, co).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        ws.Cell(++ro, co).Value = "WrapText = true";
        ws.Cell(ro, co).Style.Alignment.WrapText = true;

        workbook.SaveAs(filePath);
    }
}
