namespace XlsxSharp.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// A marker interface for types inside <c><![CDATA[<xsd:simpleType>]]></c>.
/// </summary>
public interface ISimpleType
{
    public string Name { get; }
}
