using System.Collections;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class SimpleNullableTypeReaderTests
{
    private readonly int?[] data = [1, 2, null];

    [Test]
    [MethodDataSource(nameof(SimpleNullableSourceNames))]
    public void CanGetPropertyName(IEnumerable data, string expected)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        ClassicAssert.AreEqual(expected, reader.GetPropertyName(0));
    }

    public static IEnumerable<(IEnumerable, string)> SimpleNullableSourceNames()
    {
        yield return (new int?[] { 1, 2, null }, "Int32");
        yield return (new List<double?> { 1.0, 2.0, null }, "Double");
        yield return (new decimal?[] { 1.0m, 2.0m, null }, "Decimal");
        yield return (new char?[] { 'A', 'B', null }, "Char");
        yield return (new DateTime?[] { new DateTime(2020, 1, 1), null }, "DateTime");
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        ClassicAssert.AreEqual(1, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        ClassicAssert.AreEqual(3, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        ClassicAssert.AreEqual(1, result.First().Single());
        ClassicAssert.AreEqual(Blank.Value, result.Last().Single());
    }
}
