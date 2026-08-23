using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class ArrayTypeReaderTests
{
    private readonly int[][] data = new int[][] { [1, 2, 3], [4, 5, 6] };

    [Test]
    public void GetPropertyNameReturnsNull()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        ClassicAssert.IsNull(reader.GetPropertyName(0));
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        ClassicAssert.AreEqual(3, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        ClassicAssert.AreEqual(2, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        ClassicAssert.AreEqual(1, result.First().First());
        ClassicAssert.AreEqual(3, result.First().Last());
        ClassicAssert.AreEqual(4, result.Last().First());
        ClassicAssert.AreEqual(6, result.Last().Last());
    }
}
