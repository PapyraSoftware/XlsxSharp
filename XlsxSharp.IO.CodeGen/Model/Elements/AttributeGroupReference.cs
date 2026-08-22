using System.Collections.Generic;
using XlsxSharp.IO.CodeGen.Model.TopLevel;

namespace XlsxSharp.IO.CodeGen.Model.Elements;

/// <summary>
/// An <c><![CDATA[<xsd:attributeGroup>]]></c> inside a <c><![CDATA[<xsd:complexType>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:complexType name="CT_pivotTableDefinition">
///   <xsd:attribute name="dataPosition" type="xsd:unsignedInt" use="optional"/>
///   <xsd:attributeGroup ref="AG_AutoFormat"/>
/// </xsd:complexType>
/// ]]></code>
/// </example>
/// </summary>
public class AttributeGroupReference : ILeafElement
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// Name of referenced attribute group (<see cref="AttributeGroupDefinition.Name"/>).
    /// </summary>
    public required string RefName { get; init; }
}
