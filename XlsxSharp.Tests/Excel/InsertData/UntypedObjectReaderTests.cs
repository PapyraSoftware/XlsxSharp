using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.InsertData;
using XlsxSharp.Tests.Excel.Tables;

namespace XlsxSharp.Tests.Excel.InsertData;

public class UntypedObjectReaderTests
{
    private readonly ArrayList data = new(
        new object[]
        {
            null,
            new TablesTests.TestObjectWithAttributes
            {
                Column1 = "Value 1",
                Column2 = "Value 2",
                UnOrderedColumn = 3,
                MyField = 4,
            },
            null,
            null,
            null,
            new int[] { 1, 2, 3 },
            new int[] { 4, 5, 6, 7 },
            "Separator",
            new TablesTests.TestObjectWithoutAttributes
            {
                Column1 = "Value 9",
                Column2 = "Value 10",
            },
        }
    );

    [TestCase(0, "FirstColumn")]
    [TestCase(1, "SecondColumn")]
    [TestCase(2, "SomeFieldNotProperty")]
    [TestCase(3, "UnOrderedColumn")]
    public void CanGetPropertyName(int propertyIndex, string expectedPropertyName)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        string? actualPropertyName = reader.GetPropertyName(propertyIndex);
        Assert.AreEqual(expectedPropertyName, actualPropertyName);
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(4, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(9, reader.GetRecords().Count());
    }

    [Test]
    public void CanGetData()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);

        IEnumerable<XLCellValue>[] result = [.. reader.GetRecords()];

        Assert.AreEqual(new XLCellValue[] { Blank.Value }, result[0]);
        Assert.AreEqual(new XLCellValue[] { "Value 2", "Value 1", 4, 3 }, result[1]);
        Assert.AreEqual(new XLCellValue[] { Blank.Value }, result[2]);
        Assert.AreEqual(new XLCellValue[] { Blank.Value }, result[3]);
        Assert.AreEqual(new XLCellValue[] { Blank.Value }, result[4]);
        Assert.AreEqual(new XLCellValue[] { 1, 2, 3 }, result[5]);
        Assert.AreEqual(new XLCellValue[] { 4, 5, 6, 7 }, result[6]);
        Assert.AreEqual(new XLCellValue[] { "Separator" }, result[7]);
        Assert.AreEqual(new XLCellValue[] { "Value 9", "Value 10" }, result[8]);
    }
}
