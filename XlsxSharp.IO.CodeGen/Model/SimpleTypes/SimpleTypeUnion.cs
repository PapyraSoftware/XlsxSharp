namespace XlsxSharp.IO.CodeGen.Model.SimpleTypes;

public class SimpleTypeUnion : ISimpleType
{
    public required string Name { get; init; }

    public required List<Restriction> RestrictionsUnion { get; init; } = [];
}
