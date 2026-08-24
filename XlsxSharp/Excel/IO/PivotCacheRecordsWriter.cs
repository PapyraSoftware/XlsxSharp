using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.IO.OpenXmlConst;

namespace XlsxSharp.Excel.IO;

internal class PivotCacheRecordsWriter
{
    internal static void WriteContent(
        PivotTableCacheRecordsPart recordsPart,
        XLPivotCache pivotCache
    )
    {
        XmlWriterSettings settings = new() { Encoding = XlsxSharp.XLHelper.NoBomUTF8 };

        using Stream partStream = recordsPart.GetStream(FileMode.Create);
        using XmlWriter xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartDocument();
        xml.WriteStartElement("pivotCacheRecords", Main2006SsNs);
        xml.WriteAttributeString("xmlns", "r", null, RelationshipsNs);
        xml.WriteAttributeString("xmlns", "mc", null, MarkupCompatibilityNs);

        // Mark revision as ignorable extension
        xml.WriteAttributeString("mc", "Ignorable", null, "xr");
        xml.WriteAttributeString("xmlns", "xr", null, RevisionNs);

        int recordCount = pivotCache.RecordCount;
        int fieldCount = pivotCache.FieldCount;
        for (int recordIdx = 0; recordIdx < recordCount; ++recordIdx)
        {
            xml.WriteStartElement("r");
            for (int fieldIdx = 0; fieldIdx < fieldCount; ++fieldIdx)
            {
                XLPivotCacheValues fieldValues = pivotCache.GetFieldValues(fieldIdx);
                XLPivotCacheValue value = fieldValues.GetValue(recordIdx);
                switch (value.Type)
                {
                    case XLPivotCacheValueType.Missing:
                        xml.WriteEmptyElement("m");
                        break;
                    case XLPivotCacheValueType.Number:
                        xml.WriteStartElement("n");
                        xml.WriteAttribute("v", value.GetNumber());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.Boolean:
                        xml.WriteStartElement("b");
                        xml.WriteAttribute("v", value.GetBoolean());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.Error:
                        xml.WriteStartElement("b");
                        xml.WriteAttribute("v", value.GetError().ToDisplayString());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.String:
                        xml.WriteStartElement("s");
                        xml.WriteAttribute("v", fieldValues.GetText(value));
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.DateTime:
                        xml.WriteStartElement("d");
                        xml.WriteAttribute("v", value.GetDateTime());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.Index:
                        xml.WriteStartElement("x");
                        xml.WriteAttribute("v", value.GetIndex());
                        xml.WriteEndElement();
                        break;
                    default:
                        throw new NotSupportedException();
                }
            }
            xml.WriteEndElement(); // r
        }

        xml.WriteEndElement(); // pivotCacheRecords
    }
}
