using System;
using System.IO;
using XlsxSharp.Examples.Creating;
using XlsxSharp.Examples.Loading;

namespace XlsxSharp.Examples;

public class Program
{
    public static string BaseCreatedDirectory
    {
        get
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }
    }

    public static string BaseModifiedDirectory
    {
        get
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modified");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }
    }

    private static void Main(string[] args)
    {
        CreateFiles.CreateAllFiles();
        LoadFiles.LoadAllFiles();
    }
}
