using System.Collections.Generic;
using System.Data;
using System.Linq;
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
        ClassicAssert.AreEqual("Last name", reader.GetPropertyName(0));
        ClassicAssert.AreEqual("First name", reader.GetPropertyName(1));
        ClassicAssert.AreEqual("Age", reader.GetPropertyName(2));
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
    public void CanReadValue()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        ClassicAssert.AreEqual("Smith", result.First().First());
        ClassicAssert.AreEqual(33, result.First().Last());
        ClassicAssert.AreEqual("Ivanova", result.Last().First());
        ClassicAssert.AreEqual(25, result.Last().Last());
    }
}
