using System.Diagnostics;

namespace XlsxSharp.IO.CodeGen.Model;

/// <summary>
/// <![CDATA[<xsd:attribute>]]> inside <![CDATA[<xsd:complexType>]]> or <![CDATA[<xsd:attributeGroup>]]>
/// <example>
/// <code><![CDATA[
/// <xsd:attribute name="level" type="xsd:unsignedInt" use="optional" default="0"/>
/// ]]></code>
/// </example>
/// </summary>
public class AttributeElement
{
    /// <summary>
    /// Name is technically optional in ref attribute:
    /// <code>
    ///   <![CDATA[<xsd:attribute ref="r:id" use="optional"/>]]>
    /// </code>
    /// </summary>
    public required string? Name { get; set; }

    public required string? RefName { get; set; }

    public required string? Type { get; set; }

    public AttributeUseType Use { get; set; }

    public string? DefaultValue { get; set; }

    internal bool IsOptional => this.Use is AttributeUseType.Default or AttributeUseType.Optional;

    private bool CanBeNull => this.IsOptional && this.DefaultValue is null;

    internal Variable Generate(CodeBuilder code)
    {
        Debug.Assert(this.Name is not null);
        Debug.Assert(this.Type is not null);
        code.WriteIndent()
            .Append("var ")
            .AppendVariable(this.Name)
            .Append(" = ")
            .AppendSimpleTypeMethod(this);
        if (this.DefaultValue is not null)
        {
            code.Append(" ?? ").AppendValue(this.Type, this.DefaultValue);
        }

        code.Append(";").EndLine();

        string csType = this.CanBeNull
            ? code.GetSimpleType(this.Type) + '?'
            : code.GetSimpleType(this.Type);
        return new Variable(csType, this.Name);
    }
}
