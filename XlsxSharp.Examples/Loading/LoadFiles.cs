using System;
using System.IO;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Loading;

public class LoadFiles
{
    public static void LoadAllFiles()
    {
        foreach (string file in Directory.GetFiles(Program.BaseCreatedDirectory))
        {
            FileInfo fileInfo = new(file);
            string fileName = fileInfo.Name;
            LoadAndSaveFile(
                Path.Combine(Program.BaseCreatedDirectory, fileName),
                Path.Combine(Program.BaseModifiedDirectory, fileName)
            );
        }
    }

    private static void LoadAndSaveFile(String input, String output)
    {
        XLWorkbook wb = new(input);
        wb.SaveAs(output);
        wb.SaveAs(output);
    }
}
