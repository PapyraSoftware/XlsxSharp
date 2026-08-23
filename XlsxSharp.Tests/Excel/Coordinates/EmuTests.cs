using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

internal class EmuTests
{
    [Test]
    [Arguments(0.14, AbsLengthUnit.Inch, 128_016)]
    [Arguments(2.43, AbsLengthUnit.Centimeter, 874_800)]
    [Arguments(748, AbsLengthUnit.Millimeter, 26_928_000)]
    [Arguments(23.9, AbsLengthUnit.Point, 303_530)]
    [Arguments(4.157, AbsLengthUnit.Pica, 633_527)]
    [Arguments(14.6, AbsLengthUnit.Emu, 15)]
    [Arguments(2348.52, AbsLengthUnit.Inch, null)]
    public void From_converts_value_to_emu(double value, AbsLengthUnit unit, int? emu) =>
        ClassicAssert.AreEqual(emu, Emu.From(value, unit)?.Value);

    [Test]
    [Arguments(AbsLengthUnit.Inch, 5.9912904636920388)]
    [Arguments(AbsLengthUnit.Centimeter, 15.217877777777778)]
    [Arguments(AbsLengthUnit.Millimeter, 152.17877777777778)]
    [Arguments(AbsLengthUnit.Point, 431.3729133858268)]
    [Arguments(AbsLengthUnit.Pica, 35.94774278215223)]
    [Arguments(AbsLengthUnit.Emu, 5_478_436)]
    public void To_converts_to_specified_unit(AbsLengthUnit unit, double value) =>
        ClassicAssert.AreEqual(value, Emu.From(5_478_436, AbsLengthUnit.Emu)?.To(unit));

    [Test]
    [Culture("cs-CZ")]
    public void ToString_uses_culture_invariant_format() =>
        ClassicAssert.AreEqual("1.4mm", Emu.From(1.4, AbsLengthUnit.Millimeter).ToString());
}
