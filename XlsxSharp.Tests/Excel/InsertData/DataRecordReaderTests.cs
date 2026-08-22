using System.Collections.Generic;
using System.Data;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class DataRecordReaderTests
{
    private readonly DataTable data;

    public DataRecordReaderTests()
    {
        this.data = new DataTable();
        this.data.Columns.Add("StringValue");
        this.data.Columns.Add("NumericValue", typeof(int));

        this.data.Rows.Add("Value 1", 100);
        this.data.Rows.Add("Value 2", 200);
        this.data.Rows.Add("Value 3", 300);
    }

    /// <summary>
    /// Yields the very same reader instance for every row, the way an ADO.NET data reader does. The
    /// reader under test has to materialise each record before advancing, and that is what this
    /// covers - the source of the records does not matter, only that they are <see cref="IDataRecord"/>.
    /// </summary>
    private IEnumerable<IDataRecord> GetData()
    {
        using IDataReader reader = this.data.CreateDataReader();
        while (reader.Read())
        {
            yield return reader;
        }
    }

    [Test]
    public void CanGetPropertyName()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.GetData());
        Assert.AreEqual("StringValue", reader.GetPropertyName(0));
        Assert.AreEqual("NumericValue", reader.GetPropertyName(1));
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.GetData());
        Assert.AreEqual(2, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.GetData());
        Assert.AreEqual(3, reader.GetRecords().Count());
    }

    [Test]
    public void CanGetData()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.GetData());
        IEnumerable<XLCellValue>[] result = [.. reader.GetRecords()];

        Assert.AreEqual("Value 1", result.First().First());
        Assert.AreEqual(100, result.First().Last());
        Assert.AreEqual("Value 3", result.Last().First());
        Assert.AreEqual(300, result.Last().Last());
    }
}
