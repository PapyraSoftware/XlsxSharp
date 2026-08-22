using NUnit.Framework;
using XlsxSharp.Examples.PivotTables;

namespace XlsxSharp.Tests.Examples;

[TestFixture]
public class PivotTableTests
{
    [Test]
    public void PivotTables() =>
        TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx");
}
