#nullable disable

using DocumentFormat.OpenXml.Office2010.Excel;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFConvertersExtension
{
    private static readonly Dictionary<XLConditionalFormatType, IXLCFConverterExtension> Converters;

    static XLCFConvertersExtension() =>
        XLCFConvertersExtension.Converters = new Dictionary<
            XLConditionalFormatType,
            IXLCFConverterExtension
        >()
        {
            { XLConditionalFormatType.DataBar, new XLCFDataBarConverterExtension() },
        };

    public XLCFConvertersExtension() { }

    public static ConditionalFormattingRule Convert(
        IXLConditionalFormat conditionalFormat,
        XLWorkbook.SaveContext context
    ) =>
        XLCFConvertersExtension
            .Converters[conditionalFormat.ConditionalFormatType]
            .Convert(conditionalFormat, context);
}
