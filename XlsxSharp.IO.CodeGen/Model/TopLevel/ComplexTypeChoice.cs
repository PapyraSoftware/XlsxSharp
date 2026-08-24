using XlsxSharp.IO.CodeGen.Model.Elements;

namespace XlsxSharp.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> that has <c><![CDATA[<xsd:choice>]]></c> as an element.
/// The type is inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:complexType name="CT_Tables">
///   <xsd:choice minOccurs="1" maxOccurs="unbounded">
///     <xsd:element name="m" type="CT_TableMissing"/>
///     <xsd:element name="s" type="CT_XStringElement"/>
///   </xsd:choice>
///   <xsd:attribute name="count" use="optional" type="xsd:unsignedInt"/>
/// </xsd:complexType>
/// ]]></code>
/// </example>
/// </summary>
public class ComplexTypeChoice : ComplexType
{
    public required Choice Choice { get; init; }

    internal override List<Variable> GenerateParseMethod(CodeBuilder code)
    {
        ElementsCount choicesCount = this.Choice.DetermineChoicesCount();
        switch (choicesCount)
        {
            case ElementsCount.ZeroToOne:
            {
                List<Variable> variables = this.Choice.GenerateParseContent(
                    this.Name,
                    choicesCount,
                    code,
                    true
                );
                return variables;
            }
            case ElementsCount.ZeroToMany:
            {
                List<Variable> variables = this.Choice.GenerateParseContent(
                    this.Name,
                    choicesCount,
                    code,
                    true
                );
                return variables;
            }
            case ElementsCount.OneToOne:
            {
                List<Variable> variables = this.Choice.GenerateParseContent(
                    this.Name,
                    choicesCount,
                    code,
                    true
                );
                return variables;
            }
            case ElementsCount.OneToMany:
            {
                List<Variable> variables = this.Choice.GenerateParseContent(
                    this.Name,
                    choicesCount,
                    code,
                    true
                );
                return variables;
            }
            default:
                throw new NotImplementedException();
        }
    }
}
