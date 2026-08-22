using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class SimpleNullableTypeReaderTests
{
    private readonly int?[] data = [1, 2, null];

    [TestCaseSource(nameof(SimpleNullableSourceNames))]
    public string CanGetPropertyName<T>(IEnumerable<T> data)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        return reader.GetPropertyName(0);
    }

    private static IEnumerable<TestCaseData> SimpleNullableSourceNames
    {
        get
        {
            yield return new TestCaseData(new int?[] { 1, 2, null }).Returns("Int32");
            yield return new TestCaseData(new List<double?> { 1.0, 2.0, null }).Returns("Double");
            yield return new TestCaseData(new decimal?[] { 1.0m, 2.0m, null }).Returns("Decimal");
            yield return new TestCaseData(new char?[] { 'A', 'B', null }).Returns("Char");
            yield return new TestCaseData(
                new DateTime?[] { new DateTime(2020, 1, 1), null }
            ).Returns("DateTime");
        }
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(1, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(3, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        Assert.AreEqual(1, result.First().Single());
        Assert.AreEqual(Blank.Value, result.Last().Single());
    }
}
