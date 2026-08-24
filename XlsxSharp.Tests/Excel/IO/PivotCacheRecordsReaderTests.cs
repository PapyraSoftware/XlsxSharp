using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.IO;
using XlsxSharp.IO;
using PivotCacheRecordsReader = XlsxSharp.Excel.IO.PivotCacheRecordsReader;

namespace XlsxSharp.Tests.Excel.IO;

internal class PivotCacheRecordsReaderTests
{
    [Test]
    public void Can_read_all_record_item_types()
    {
        XLPivotCacheSharedItems sharedItems = new();
        sharedItems.Add("First shared item");
        sharedItems.Add("Second shared item");

        ReadRecords(
            new[] { "Field 1" },
            $"""
            <pivotCacheRecords xmlns="{OpenXmlConst.Main2006SsNs}">
              <r>
                <m/>
              </r>
              <r>
                <n v="5.5"/>
              </r>
              <r>
                <b v="true"/>
              </r>
              <r>
                <e v="#NUM!"/>
              </r>
              <r>
                <s v="Text"/>
              </r>
              <r>
                <d v="2020-10-05"/>
              </r>
              <r>
                <x v="1"/>
              </r>
            </pivotCacheRecords>
            """,
            (cache, reader) =>
            {
                reader.ReadRecordsToCache();

                XLPivotCacheValues values = cache.GetFieldValues(0);
                CollectionAssert.AreEquivalent(
                    new XLCellValue[]
                    {
                        Blank.Value,
                        5.5,
                        true,
                        XLError.NumberInvalid,
                        "Text",
                        new DateTime(2020, 10, 5),
                        "Second shared item",
                    },
                    values.GetCellValues()
                );
            },
            sharedItems
        );
    }

    [Test]
    [Arguments("<m/>")]
    [Arguments("<m/><m/><m/>")]
    public void All_records_must_have_same_number_of_items_as_there_is_cache_fields(
        string recordItems
    ) =>
        ReadRecords(
            new[] { "Field 1", "Field 2" },
            $"""
            <pivotCacheRecords xmlns="{OpenXmlConst.Main2006SsNs}">
              <r>{recordItems}</r>
            </pivotCacheRecords>
            """,
            (_, reader) =>
            {
                PartStructureException ex = ClassicAssert.Throws<PartStructureException>(
                    reader.ReadRecordsToCache
                );
                StringAssert.StartsWith(
                    PartStructureException.IncorrectElementsCount().Message,
                    ex.Message
                );
            }
        );

    private static void ReadRecords(
        IReadOnlyList<string> fieldNames,
        string recordsXml,
        Action<XLPivotCache, PivotCacheRecordsReader> assert,
        XLPivotCacheSharedItems sharedItems = null
    )
    {
        using XLWorkbook wb = new();
        XLPivotCache cache = wb.PivotCachesInternal.Add(new XLPivotSourceConnection(0));
        sharedItems ??= new XLPivotCacheSharedItems();
        foreach (string fieldName in fieldNames)
        {
            cache.AddCachedField(
                fieldName,
                new XLPivotCacheValues(sharedItems, new XLPivotCacheValuesStats())
            );
        }

        using MemoryStream stream = new(XLHelper.NoBomUTF8.GetBytes(recordsXml));
        using XmlTreeReader xmlTreeReader = new(stream, XmlToEnumMapper.Instance, true);
        PivotCacheRecordsReader reader = new(xmlTreeReader, cache);
        assert(cache, reader);
    }
}
