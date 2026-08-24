namespace XlsxSharp.Parser.Tests;

public class FormulaConverterToR1C1Tests
{
    [Test]
    [Arguments("true", "true")]
    [Arguments("FALSE", "FALSE")]
    [Arguments("1", "1")]
    [Arguments("\"Text\"", "\"Text\"")]
    [Arguments("\"\"", "\"\"")]
    [Arguments("#DIV/0!", "#DIV/0!")]
    public async Task Constants(string a1, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, 1, 1)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments(" { 1 }   ", " { 1 }   ")]
    [Arguments("{1,2}", "{1,2}")]
    [Arguments("{1;2}", "{1;2}")]
    [Arguments("{ 1, 2;  3, 4 }", "{ 1, 2;  3, 4 }")]
    [Arguments("{TRUE}", "{TRUE}")]
    [Arguments("{ #DIV/0! } ", "{ #DIV/0! } ")]
    [Arguments("{ \"\"}", "{ \"\"}")]
    [Arguments("{ \"Hello world\" }", "{ \"Hello world\" }")]
    public async Task Array(string a1, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, 1, 1)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("SUM(#REF!$B$7)", "SUM(#REF!)")]
    [Arguments("SUM(Sheet!#REF!)", "SUM(Sheet!#REF!)")]
    [Arguments("SUM(#REF!#REF!)", "SUM(#REF!)")]
    public async Task ErrorNode(string a1, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, 1, 1)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("B4", 4, 2, "RC")] // A1 Relative
    [Arguments("B4", 5, 2, "R[-1]C")]
    [Arguments("B4", 4, 1, "RC[1]")]
    [Arguments("B4", 2, 1, "R[2]C[1]")]
    [Arguments("$B4", 4, 2, "RC2")] // A1 Mixed
    [Arguments("B$4", 4, 2, "R4C")]
    [Arguments("$B$4", 9, 7, "R4C2")] // A1 Absolute
    [Arguments("B4:B4", 2, 1, "R[2]C[1]")] // Both are same
    [Arguments("B4:B$4", 2, 1, "R[2]C[1]:R4C[1]")]
    [Arguments("C5:Z14", 2, 6, "R[3]C[-3]:R[12]C[20]")]
    public async Task Reference(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("'Sheet name'!$D6", 4, 1, "'Sheet name'!R[2]C4")]
    [Arguments("January!$D2", 4, 1, "January!R[-2]C4")]
    [Arguments("'A''B'!$D2", 4, 1, "'A''B'!R[-2]C4")]
    public async Task SheetReference(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("January:December!$D6", 4, 1, "January:December!R[2]C4")]
    [Arguments("'Johnny''s:Denny''s'!$D6", 4, 1, "'Johnny''s:Denny''s'!R[2]C4")]
    public async Task Reference3D(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[74]Sheet5!$D6", 4, 1, "[74]Sheet5!R[2]C4")]
    [Arguments("'[6]Johnny''s house'!$D6", 4, 1, "'[6]Johnny''s house'!R[2]C4")]
    public async Task ExternalSheetReference(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[74]Sheet1:Sheet7!$D6", 4, 1, "[74]Sheet1:Sheet7!R[2]C4")]
    [Arguments("'[6]Johnny''s:Danny''s'!$D6", 4, 1, "'[6]Johnny''s:Danny''s'!R[2]C4")]
    public async Task ExternalReference3D(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("RAND()", 4, 1, "RAND()")]
    [Arguments("RAND(   )  ", 4, 1, "RAND(   )  ")]
    [Arguments("SIN( F8 ) ", 2, 3, "SIN( R[6]C[3] ) ")]
    [Arguments("MOD(F8 ,  $A$1)", 2, 3, "MOD(R[6]C[3] ,  R1C1)")]
    [Arguments("IF(TRUE,,)", 2, 3, "IF(TRUE,,)")]
    public async Task Function(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("Sheet1!UDF.SHEET.FUNC($F$1)", 4, 1, "Sheet1!UDF.SHEET.FUNC(R1C6)")]
    [Arguments("'Johnny''s'!UDF.SHEET.FUNC($F$1)", 4, 1, "'Johnny''s'!UDF.SHEET.FUNC(R1C6)")]
    [Arguments("Sheet1!UDF.SHEET.FUNC( $F$1  ) ", 4, 1, "Sheet1!UDF.SHEET.FUNC( R1C6  ) ")]
    public async Task SheetFunction(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[4]!UDF.SHEET.FUNC($F$1)", 4, 1, "[4]!UDF.SHEET.FUNC(R1C6)")]
    [Arguments("[4]!UDF.SHEET.FUNC( $F$1  )", 4, 1, "[4]!UDF.SHEET.FUNC( R1C6  )")]
    public async Task ExternalFunction(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[4]Sheet1!UDF.SHEET.FUNC($F$1)", 4, 1, "[4]Sheet1!UDF.SHEET.FUNC(R1C6)")]
    [Arguments("'[7]Johnny''s'!UDF.SHEET.FUNC($F$1)", 4, 1, "'[7]Johnny''s'!UDF.SHEET.FUNC(R1C6)")]
    [Arguments("[2]Sheet1!F(  ) ", 4, 1, "[2]Sheet1!F(  ) ")]
    [Arguments("[2]Sheet1!F(  $A$2,$F$1 ) ", 4, 1, "[2]Sheet1!F(  R2C1,R1C6 ) ")]
    public async Task ExternalSheetFunction(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("$A$5(TRUE,7)", 10, 15, "R5C1(TRUE,7)")]
    [Arguments("$A$5(TRUE ,  7)", 10, 15, "R5C1(TRUE ,  7)")]
    public async Task CellFunction(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[]", 10, 15, "[]")]
    [Arguments("[#Headers]", 10, 15, "[#Headers]")]
    [Arguments("[#Data]", 10, 15, "[#Data]")]
    [Arguments("[#Totals]", 10, 15, "[#Totals]")]
    [Arguments("[#All]", 10, 15, "[#All]")]
    [Arguments("[#This Row]", 10, 15, "[#This Row]")]
    [Arguments("[[#Headers],[#Data]]", 10, 15, "[[#Headers],[#Data]]", Skip = "Parser fail")]
    [Arguments("[[#Data],[#Totals]]", 10, 15, "[[#Data],[#Totals]]", Skip = "Parser fail")]
    [Arguments("[Column]", 10, 15, "[Column]")]
    [Arguments("[Space column]", 10, 15, "[Space column]")]
    [Arguments("[[#Data],[Column]]", 10, 15, "[[#Data],[Column]]")]
    [Arguments("[[#Data],[Space column]]", 10, 15, "[[#Data],[Space column]]")]
    [Arguments("[[#Headers],[#Data],[Column]]", 10, 15, "[[#Headers],[#Data],[Column]]")]
    [Arguments("[[#Data],[#Totals],[Space column]]", 10, 15, "[[#Data],[#Totals],[Space column]]")]
    [Arguments("[[#All],[Column 1]:[Column 2]]", 10, 15, "[[#All],[Column 1]:[Column 2]]")]
    [Arguments(
        "[[#Data],[#Totals],[Column 1]:[Column 2]]",
        10,
        15,
        "[[#Data],[#Totals],[Column 1]:[Column 2]]"
    )]
    public async Task StructureReference(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments(
        "Table1[[#Data],[#Totals],[Column 1]:[Column 2]]",
        10,
        15,
        "Table1[[#Data],[#Totals],[Column 1]:[Column 2]]"
    )]
    public async Task TableStructureReference(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[1]Table1[Column]", 10, 15, "[1]Table1[Column]", Skip = "Parser fail")]
    public async Task ExternalStructureReference(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments(" some_name + other_name", 1, 1, "some_name + other_name")]
    public async Task Name(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("Sheet1!defined_name", 1, 1, "Sheet1!defined_name")]
    [Arguments("'Sheet name'!defined_name", 1, 1, "'Sheet name'!defined_name")]
    public async Task SheetName(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[4]!defined_name", 1, 1, "[4]!defined_name")]
    public async Task ExternalName(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("[0]Sheet5!name", 1, 1, "[0]Sheet5!name")]
    [Arguments("'[4]Happy sheet'!data", 1, 1, "'[4]Happy sheet'!data")]
    public async Task ExternalSheetName(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("+B3", 1, 1, "+R[2]C[1]")]
    [Arguments("-8", 1, 1, "-8")]
    [Arguments(" - 8 ", 1, 1, " - 8 ")]
    [Arguments("100 %", 1, 1, "100 %")]
    [Arguments("@D8", 2, 5, "@R[6]C[-1]")]
    [Arguments("D8#", 2, 5, "R[6]C[-1]#")]
    public async Task Unary(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments("5+1", 1, 1, "5+1")]
    [Arguments("5 +  1 ", 1, 1, "5 +  1 ")]
    [Arguments("B3 +  $D$8 ", 1, 1, "R[2]C[1] +  R8C4 ")]
    [Arguments("B3 /  $D$8 ", 1, 1, "R[2]C[1] /  R8C4 ")]
    public async Task Binary(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }

    [Test]
    [Arguments(" ( 1 + 2 ) / 4", 1, 1, " ( 1 + 2 ) / 4")]
    [Arguments("(1+((3 + A4)))", 1, 1, "(1+((3 + R[3]C)))")]
    public async Task Nested(string a1, int row, int col, string r1c1)
    {
        await Assert.That(FormulaConverter.ToR1C1(a1, row, col)).IsEqualTo(r1c1);
    }
}
