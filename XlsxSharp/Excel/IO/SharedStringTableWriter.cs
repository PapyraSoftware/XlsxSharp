using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.Utils;
using static XlsxSharp.Excel.IO.OpenXmlConst;
using static XlsxSharp.Excel.XLWorkbook;

namespace XlsxSharp.Excel.IO;

internal class SharedStringTableWriter
{
    internal static void GenerateSharedStringTablePartContent(
        XLWorkbook workbook,
        SharedStringTablePart sharedStringTablePart,
        SaveContext context
    )
    {
        // Call all table headers to make sure their names are filled
        workbook.Worksheets.ForEach(w => w.Tables.ForEach(t => _ = ((XLTable)t).FieldNames.Count));

        XmlWriterSettings settings = new()
        {
            CloseOutput = true,
            Encoding = XlsxSharp.XLHelper.NoBomUTF8,
        };
        Stream partStream = sharedStringTablePart.GetStream(FileMode.Create);
        using XmlWriter xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartDocument();

        // Due to streaming and XLWorkbook structure, we don't know count before strings are written.
        // Attributes count and uniqueCount are optional thus are omitted.
        xml.WriteStartElement("x", "sst", Main2006SsNs);

        SharedStringTable sst = workbook.SharedStringTable;
        List<int> map = sst.GetConsecutiveMap();
        context.SstMap = map;
        for (int sharedStringId = 0; sharedStringId < map.Count; ++sharedStringId)
        {
            int continuousId = map[sharedStringId];
            if (continuousId < 0)
            {
                continue;
            }

            XLImmutableRichText? richText = sst.GetRichText(sharedStringId);
            if (richText is not null)
            {
                xml.WriteStartElement("si", Main2006SsNs);
                TextSerializer.WriteRichTextElements(xml, richText, context);
                xml.WriteEndElement(); // si
            }
            else
            {
                xml.WriteStartElement("si", Main2006SsNs);
                xml.WriteStartElement("t", Main2006SsNs);
                string sharedString = sst[sharedStringId];
                if (!sharedString.Trim().Equals(sharedString))
                {
                    xml.WritePreserveSpaceAttr();
                }

                xml.WriteString(XmlEncoder.EncodeString(sharedString));
                xml.WriteEndElement(); // t
                xml.WriteEndElement(); // si
            }
        }

        xml.WriteEndElement(); // SharedStringTable
        xml.Close();
    }
}
