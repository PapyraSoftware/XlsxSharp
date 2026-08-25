using System.Collections;
using XlsxSharp.Excel;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class SimpleTypeReaderTests
{
    private readonly int[] data = [1, 2, 3];

    [Test]
    [MethodDataSource(nameof(SimpleSourceNames))]
    public void CanGetPropertyName(IEnumerable data, string expected)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        ClassicAssert.AreEqual(expected, reader.GetPropertyName(0));
    }

    public static IEnumerable<(IEnumerable, string)> SimpleSourceNames()
    {
        yield return (new[] { 1, 2, 3 }, "Int32");
        yield return (new List<double> { 1.0, 2.0, 3.0 }, "Double");
        yield return (new[] { 1.0m, 2.0m, 3.0m }, "Decimal");
        yield return (new[] { "A", "B", "C" }, "String");
        yield return (new[] { 'A', 'B', 'C' }, "Char");
        yield return (new[] { new DateTime(2020, 1, 1) }, "DateTime");
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
        ClassicAssert.AreEqual(3, result.Last().Single());
    }
}
