using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
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
        Assert.IsNull(reader.GetPropertyName(0));
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(3, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(2, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        Assert.AreEqual(1, result.First().First());
        Assert.AreEqual(3, result.First().Last());
        Assert.AreEqual(4, result.Last().First());
        Assert.AreEqual(6, result.Last().Last());
    }
}
