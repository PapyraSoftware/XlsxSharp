using System.IO.Compression;
using System.Text;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Comments;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Drawings.Style;

namespace XlsxSharp.Tests.Excel.Comments;

/// <summary>
/// Pins the VML that carries comment shapes. None of the reference workbooks contain a comment,
/// so the byte comparisons that guard every other part do not cover this one, and Excel is
/// unforgiving about VML. This test is the missing net: it records what the writer produces so
/// that a change to it has to be deliberate.
/// </summary>
public class CommentVmlOutputTests
{
    [Test]
    public void CommentVmlIsWrittenAsRecorded()
    {
        string vml = SaveAndReadVml(worksheet =>
        {
            worksheet.Cell("B2").CreateComment().AddText("Plain");

            IXLComment styled = worksheet.Cell("D5").CreateComment();
            styled.AddText("Styled");
            styled.Style.ColorsAndLines.FillColor = XLColor.LightYellow;
            styled.Style.ColorsAndLines.LineColor = XLColor.Red;
            styled.Style.ColorsAndLines.LineWeight = 2.5;
            styled.Style.ColorsAndLines.LineDash = XLDashStyle.RoundDot;
            styled.Style.Alignment.Horizontal = XLDrawingHorizontalAlignment.Center;
            styled.Style.Alignment.Vertical = XLDrawingVerticalAlignment.Bottom;
            styled.Style.Protection.Locked = false;
            styled.Style.Protection.LockText = false;
            styled.SetVisible();
        });

        ClassicAssert.AreEqual(ExpectedVml, vml);
    }

    /// <summary>
    /// Recorded from the writer. Regenerate deliberately, never to make a red test green: a
    /// change here is a change to what Excel gets handed.
    /// </summary>
    private const string ExpectedVml = """
        <xml><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><v:stroke joinstyle="miter" /><v:path gradientshapeok="true" o:connecttype="rect" /></v:shapetype><v:shape id="_x0000_s1" style="position:absolute; visibility:hidden;width:144pt;height:59.25pt;z-index:1" o:insetmode="auto" fillcolor="#FFFFE1" strokecolor="#000000" strokeweight="0.75pt" type="#_x0000_t202" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><v:fill color2="#FFFFE1" /><v:stroke linestyle="single" dashstyle="solid" /><v:shadow obscured="true" color="black" /><v:path o:connecttype="none" /><v:textbox /><xvml:ClientData ObjectType="Note" xmlns:xvml="urn:schemas-microsoft-com:office:excel"><xvml:MoveWithCells>True</xvml:MoveWithCells><xvml:SizeWithCells>True</xvml:SizeWithCells><xvml:Anchor>2, 15, 0, 8, 4, 33, 4, 7</xvml:Anchor><xvml:TextHAlign>left</xvml:TextHAlign><xvml:TextVAlign>top</xvml:TextVAlign><xvml:AutoFill>False</xvml:AutoFill><xvml:Row>1</xvml:Row><xvml:Column>1</xvml:Column><xvml:Locked>True</xvml:Locked><xvml:LockText>True</xvml:LockText><xvml:Visible>False</xvml:Visible></xvml:ClientData></v:shape><v:shape id="_x0000_s2" style="position:absolute; visibility:visible;width:144pt;height:59.25pt;z-index:2" o:insetmode="auto" fillcolor="#FFFFE0" strokecolor="#FF0000" strokeweight="2.5pt" type="#_x0000_t202" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><v:fill color2="#FFFFE0" /><v:stroke linestyle="single" endcap="round" dashstyle="shortDot" /><v:shadow obscured="true" color="black" /><v:path o:connecttype="none" /><v:textbox /><xvml:ClientData ObjectType="Note" xmlns:xvml="urn:schemas-microsoft-com:office:excel"><xvml:MoveWithCells>True</xvml:MoveWithCells><xvml:SizeWithCells>True</xvml:SizeWithCells><xvml:Anchor>4, 15, 3, 8, 6, 33, 7, 7</xvml:Anchor><xvml:TextHAlign>center</xvml:TextHAlign><xvml:TextVAlign>bottom</xvml:TextVAlign><xvml:AutoFill>False</xvml:AutoFill><xvml:Row>4</xvml:Row><xvml:Column>3</xvml:Column><xvml:Locked>False</xvml:Locked><xvml:LockText>False</xvml:LockText><xvml:Visible>True</xvml:Visible></xvml:ClientData></v:shape></xml>
        """;

    private static string SaveAndReadVml(Action<IXLWorksheet> build)
    {
        using MemoryStream stream = new();
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet worksheet = workbook.AddWorksheet("Comments");
            build(worksheet);
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        using ZipArchive archive = new(stream, ZipArchiveMode.Read);
        ZipArchiveEntry entry =
            archive.Entries.FirstOrDefault(e =>
                e.FullName.Contains("vmldrawing", StringComparison.OrdinalIgnoreCase)
            )
            ?? throw new InvalidOperationException(
                "The saved workbook has no VML drawing part. Entries: "
                    + string.Join(", ", archive.Entries.Select(e => e.FullName))
            );

        using Stream entryStream = entry.Open();
        using StreamReader reader = new(entryStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
