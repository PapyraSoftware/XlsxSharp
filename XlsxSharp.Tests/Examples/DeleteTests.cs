using XlsxSharp.Examples.Delete;

namespace XlsxSharp.Tests.Examples;

public class DeleteTests
{
    [Test]
    public void DeleteFewWorksheets() =>
        TestHelper.RunTestExample<DeleteFewWorksheets>(@"Delete\DeleteFewWorksheets.xlsx");

    [Test]
    public void RemoveRows() => TestHelper.RunTestExample<DeleteRows>(@"Delete\RemoveRows.xlsx");
}
