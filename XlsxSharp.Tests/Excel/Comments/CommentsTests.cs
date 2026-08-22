using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Drawings.Style;
using XlsxSharp.Extensions;
using Point = System.Drawing.Point;

namespace XlsxSharp.Tests.Excel.Comments;

public class CommentsTests
{
    [Test]
    public void CanConvertVmlPaletteEntriesToColors()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\CommentsWithColorNamesAndIndexes.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            IXLCell c = ws.FirstCellUsed();

            // None indicates an absence of a color
            XLColor lineColor = c.GetComment().Style.ColorsAndLines.LineColor;
            Assert.AreEqual(XLColorType.Color, lineColor.ColorType);
            Assert.AreEqual("00000000", lineColor.Color.ToHex());

            XLColor bgColor = c.GetComment().Style.ColorsAndLines.FillColor;
            Assert.AreEqual(XLColorType.Color, bgColor.ColorType);
            Assert.AreEqual("FFFFFFE1", bgColor.Color.ToHex());
        }
    }

    [Test]
    public void CopyCommentStyle()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            string strExcelComment =
                "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "2) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "3) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "4) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "5) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "6) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "7) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "8) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;
            strExcelComment =
                strExcelComment
                + "9) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC"
                + Environment.NewLine;

            IXLCell cell = ws.Cell(2, 2).SetValue("Comment 1");

            cell.GetComment().SetVisible(false).AddText(strExcelComment);

            cell.GetComment().Style.Alignment.SetAutomaticSize();

            cell.GetComment().Style.ColorsAndLines.SetFillColor(XLColor.Red);

            ws.Row(1).InsertRowsAbove(1);

            Action<IXLCell> validate = c =>
            {
                Assert.IsTrue(c.GetComment().Style.Alignment.AutomaticSize);
                Assert.AreEqual(XLColor.Red, c.GetComment().Style.ColorsAndLines.FillColor);
            };

            validate(ws.Cell("B3"));

            ws.Column(1).InsertColumnsBefore(2);

            validate(ws.Cell("D3"));

            ws.Column(1).Delete();

            validate(ws.Cell("C3"));

            ws.Row(1).Delete();

            validate(ws.Cell("C2"));
        }
    }

    [Test]
    public void EnsureUnaffectedCommentAndVmlPartIdsAndUris()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx")
            )
        )
        using (MemoryStream ms = new())
        {
            string commentPartId;
            string commentPartUri;

            string vmlPartId;
            string vmlPartUri;

            using (SpreadsheetDocument ssd = SpreadsheetDocument.Open(stream, isEditable: false))
            {
                WorkbookPart wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
                WorksheetPart wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

                WorksheetCommentsPart wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
                commentPartId = wsp.GetIdOfPart(wscp);
                commentPartUri = wscp.Uri.ToString();

                VmlDrawingPart vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
                vmlPartId = wsp.GetIdOfPart(vmlp);
                vmlPartUri = vmlp.Uri.ToString();
            }

            stream.Position = 0;
            stream.CopyTo(ms);
            ms.Position = 0;

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                Assert.IsTrue(ws.FirstCell().HasComment);

                wb.SaveAs(ms);
            }

            ms.Position = 0;

            using (SpreadsheetDocument ssd = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                WorkbookPart wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
                WorksheetPart wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

                WorksheetCommentsPart wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
                Assert.AreEqual(commentPartUri, wscp.Uri.ToString());
                Assert.AreEqual(commentPartId, wsp.GetIdOfPart(wscp));

                VmlDrawingPart vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
                Assert.AreEqual(vmlPartUri, vmlp.Uri.ToString());
                Assert.AreEqual(vmlPartId, wsp.GetIdOfPart(vmlp));
            }
        }
    }

    [Test]
    public void SavingDoesNotCauseTwoRootElements() // See #1157
    {
        using (MemoryStream ms = new())
        {
            using (
                Stream stream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx")
                )
            )
            using (XLWorkbook wb = new(stream))
            {
                wb.SaveAs(ms);
            }

            Assert.DoesNotThrow(() => new XLWorkbook(ms));
        }
    }

    [Test]
    public void CanLoadCommentVisibility()
    {
        using (
            Stream inputStream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Other\Drawings\Comments\inputfile.xlsx")
            )
        )
        using (XLWorkbook workbook = new(inputStream))
        {
            IXLWorksheet ws = workbook.Worksheets.First();

            Assert.True(ws.Cell("A1").GetComment().Visible);
            Assert.False(ws.Cell("A4").GetComment().Visible);
        }
    }

    [Test]
    public void Margins_are_converted_to_physical_length()
    {
        // Technically, it's insets on a textbox. Each comment uses a different unit, but all
        // should have same final dimension at left and top margin (easily visible in the
        // sheet). Tested units: in, cm, mm, pt, pc, emu, px, em, ex. Pixels are converted
        // through supplied DPI.
        // The last comment in vmlDrawing1 also has invalid units and number. These are
        // converted to 0, so we don't crash on load (Excel also ignores invalid values).
        string[] commentCells = ["A1", "A7", "A16", "A22", "A28"];
        TestHelper.LoadAndAssert(
            (_, ws) =>
            {
                foreach (string commentCell in commentCells)
                {
                    IXLCell cell = ws.Cell(commentCell);
                    Assert.True(cell.HasComment);
                    IXLDrawingMargins margins = cell.GetComment().Style.Margins;

                    Assert.AreEqual(0.5, margins.Left);
                    Assert.AreEqual(0.75, margins.Top);

                    Assert.AreEqual(0, margins.Right);
                    Assert.AreEqual(0, margins.Bottom);
                }
            },
            @"Other\Comments\InsetsUnitConversion.xlsx",
            new LoadOptions { Dpi = new Point(120, 120) }
        );
    }
}
