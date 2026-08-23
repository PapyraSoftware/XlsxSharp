using XlsxSharp.Examples.PivotTables;

namespace XlsxSharp.Tests.Examples;

public class PivotTableTests
{
    [Test]
    public void PivotTables() =>
        TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx");
}
