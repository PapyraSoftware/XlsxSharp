namespace XlsxSharp.Utils;

internal static class OpenXmlHelper
{
    internal static int NormalizeRotation(uint textRotation) =>
        textRotation switch
        {
            <= 90 => (int)textRotation,
            <= 180 => 90 - (int)textRotation,
            255 => 255,
            _ => throw new ArgumentOutOfRangeException(),
        };
}
