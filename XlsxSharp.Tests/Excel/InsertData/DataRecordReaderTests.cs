using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class DataRecordReaderTests
{
    private readonly string connectionString =
        @"Data Source=(localdb)\MSSQLLocalDB;Integrated Security=True;Connect Timeout=1";

    private IEnumerable<IDataRecord> GetData()
    {
        const string queryString =
            @"
            select 'Value 1' as StringValue, 100 as NumericValue
            union all
            select 'Value 2', 200
            union all
            select 'Value 3', 300";

        using (SqlConnection connection = new(this.connectionString))
        using (SqlCommand command = new(queryString, connection))
        {
            try
            {
                connection.Open();
            }
            catch
            {
                Assert.Ignore("Could not connect to localdb");
            }

            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    yield return reader;
                }
            }
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
