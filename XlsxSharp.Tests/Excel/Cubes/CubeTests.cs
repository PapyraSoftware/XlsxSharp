namespace XlsxSharp.Tests.Excel.Cubes;

public class CubeTests
{
    [Test]
    public void CalLoadAndSaveCubeFromRange() =>
        TestHelper.LoadSaveAndCompare(
            @"Other\Cubes\CubeFromRange-Input.xlsx",
            @"Other\Cubes\CubeFromRange-Output.xlsx"
        );
}
