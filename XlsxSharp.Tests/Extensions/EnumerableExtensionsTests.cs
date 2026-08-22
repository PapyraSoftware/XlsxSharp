using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Extensions;
using XlsxSharp.Tests.Excel.Tables;

namespace XlsxSharp.Tests.Extensions;

public class EnumerableExtensionsTests
{
    [Test]
    public void CanGetItemType()
    {
        int[] array = [];
        Assert.AreEqual(typeof(int), array.GetItemType());

        List<double> list = [];
        Assert.AreEqual(typeof(double), list.GetItemType());
        Assert.AreEqual(typeof(double), list.AsEnumerable().GetItemType());

        IEnumerable<IEnumerable> enumerable = (List<string>)[];
        Assert.AreEqual(typeof(string), enumerable.GetItemType());

        enumerable = (List<List<string>>)[];
        Assert.AreEqual(typeof(List<string>), enumerable.GetItemType());

        enumerable = (List<int[]>)[];
        Assert.AreEqual(typeof(int[]), enumerable.GetItemType());

        var anonymousIterator = new List<TablesTests.TestObjectWithoutAttributes>().Select(o => new
        {
            FirstName = o.Column1,
            LastName = o.Column2,
        });

        //expectedType can be something like <>f__AnonymousType9`2[System.String,System.String]
        //but since that `9` may differ with new anonymous types declare in the assembly
        //check the beginning and the ending of the actual type
        string expectedTypeStart = "<>f__AnonymousType";
        string expectedTypeEnd = "`2[System.String,System.String]";
        string actualType = anonymousIterator.GetItemType().ToString();
        Assert.True(actualType.StartsWith(expectedTypeStart));
        Assert.True(actualType.EndsWith(expectedTypeEnd));

        IEnumerable<object> obj = anonymousIterator;
        actualType = obj.GetItemType().ToString();
        Assert.True(actualType.StartsWith(expectedTypeStart));
        Assert.True(actualType.EndsWith(expectedTypeEnd));
    }

    [Test]
    public void SkipLastSkipsLastElementOfEnumerable()
    {
        IEnumerable<int> empty = Array.Empty<int>().SkipLast();
        CollectionAssert.IsEmpty(empty);

        IEnumerable<int> oneElement = new[] { 1 }.SkipLast();
        CollectionAssert.IsEmpty(oneElement);

        IEnumerable<int> twoElements = new[] { 1, 2 }.SkipLast();
        CollectionAssert.AreEqual(new[] { 1 }, twoElements);
    }

    [Test]
    public void WhereNotNullRemovesNullElements()
    {
        int?[] source = [1, null, 2];

        IEnumerable<int> result = source.WhereNotNull(x => x);

        CollectionAssert.AreEqual(new[] { 1, 2 }, result);
    }
}
