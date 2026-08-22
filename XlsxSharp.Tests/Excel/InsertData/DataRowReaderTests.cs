using System.Collections.Generic;
using System.Data;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class DataRowReaderTests
{
    private readonly DataTable data;

    public DataRowReaderTests()
    {
        this.data = new DataTable();
        this.data.Columns.Add("Last name");
        this.data.Columns.Add("First name");
        this.data.Columns.Add("Age", typeof(int));

        this.data.Rows.Add("Smith", "John", 33);
        this.data.Rows.Add("Ivanova", "Olga", 25);
    }

    [Test]
    public void CanGetPropertyName()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual("Last name", reader.GetPropertyName(0));
        Assert.AreEqual("First name", reader.GetPropertyName(1));
        Assert.AreEqual("Age", reader.GetPropertyName(2));
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
    public void CanReadValue()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        Assert.AreEqual("Smith", result.First().First());
        Assert.AreEqual(33, result.First().Last());
        Assert.AreEqual("Ivanova", result.Last().First());
        Assert.AreEqual(25, result.Last().Last());
    }
}
