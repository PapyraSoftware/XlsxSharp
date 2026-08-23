using XlsxSharp.Examples.Loading;

namespace XlsxSharp.Tests.Examples;

public class LoadingTests
{
    [Test]
    public void ChangingBasicTable() =>
        TestHelper.RunTestExample<ChangingBasicTable>(@"Loading\ChangingBasicTable.xlsx");
}
