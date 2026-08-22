using System.IO;
using System.Reflection;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Drawings;

namespace XlsxSharp.Examples.ImageHandling;

public class ImageFormats : IXLExample
{
    public void Create(string filePath)
    {
        using XLWorkbook wb = new();
        using (
            Stream fs = Assembly
                .GetExecutingAssembly()
                .GetManifestResourceStream("XlsxSharp.Examples.Resources.ImageHandling.jpg")
        )
        {
            #region Jpeg

            IXLWorksheet ws = wb.Worksheets.Add("Jpg");
            ws.AddPicture(fs, XLPictureFormat.Jpeg, "JpegImage").MoveTo(ws.Cell(1, 1));

            #endregion Jpeg
        }

        using (
            Stream fs = Assembly
                .GetExecutingAssembly()
                .GetManifestResourceStream("XlsxSharp.Examples.Resources.ImageHandling.png")
        )
        {
            #region Png

            IXLWorksheet ws = wb.Worksheets.Add("Png");
            ws.AddPicture(fs, XLPictureFormat.Png, "PngImage").MoveTo(ws.Cell(1, 1));

            #endregion Png

            wb.SaveAs(filePath);
        }
    }
}
