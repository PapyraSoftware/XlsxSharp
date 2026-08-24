namespace XlsxSharp.Examples;

public static class ExampleHelper
{
    public static string GetTempFilePath() => Path.GetTempFileName();

    public static string GetTempFilePath(string filePath)
    {
        string extension = Path.GetExtension(filePath);
        string tempFilePath = GetTempFilePath();
        return Path.ChangeExtension(tempFilePath, extension);
    }
}
