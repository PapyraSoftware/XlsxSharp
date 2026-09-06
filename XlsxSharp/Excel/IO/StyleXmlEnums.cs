using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// The enumerations of the font, keyed by the string OOXML writes for them. Fonts are described
/// the same way wherever they appear - in <c>xl/styles.xml</c>, in a rich text run and in a
/// comment - so these are shared rather than owned by one part.
/// </summary>
internal static class StyleXmlEnums
{
    internal static XLFontUnderlineValues ParseUnderline(string value) =>
        value switch
        {
            "single" => XLFontUnderlineValues.Single,
            "double" => XLFontUnderlineValues.Double,
            "singleAccounting" => XLFontUnderlineValues.SingleAccounting,
            "doubleAccounting" => XLFontUnderlineValues.DoubleAccounting,
            "none" => XLFontUnderlineValues.None,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLFontUnderlineValues value) =>
        value switch
        {
            XLFontUnderlineValues.Single => "single",
            XLFontUnderlineValues.Double => "double",
            XLFontUnderlineValues.SingleAccounting => "singleAccounting",
            XLFontUnderlineValues.DoubleAccounting => "doubleAccounting",
            XLFontUnderlineValues.None => "none",
            _ => throw UnknownValue(value),
        };

    internal static XLFontVerticalTextAlignmentValues ParseVerticalTextAlignment(string value) =>
        value switch
        {
            "baseline" => XLFontVerticalTextAlignmentValues.Baseline,
            "superscript" => XLFontVerticalTextAlignmentValues.Superscript,
            "subscript" => XLFontVerticalTextAlignmentValues.Subscript,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLFontVerticalTextAlignmentValues value) =>
        value switch
        {
            XLFontVerticalTextAlignmentValues.Baseline => "baseline",
            XLFontVerticalTextAlignmentValues.Superscript => "superscript",
            XLFontVerticalTextAlignmentValues.Subscript => "subscript",
            _ => throw UnknownValue(value),
        };

    internal static XLFontScheme ParseFontScheme(string value) =>
        value switch
        {
            "none" => XLFontScheme.None,
            "major" => XLFontScheme.Major,
            "minor" => XLFontScheme.Minor,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLFontScheme value) =>
        value switch
        {
            XLFontScheme.None => "none",
            XLFontScheme.Major => "major",
            XLFontScheme.Minor => "minor",
            _ => throw UnknownValue(value),
        };

    private static ArgumentOutOfRangeException UnknownValue<T>(T value)
        where T : struct, Enum => new(nameof(value), value, $"Unknown {typeof(T).Name} value.");
}
