using System;
using System.Collections.Generic;
using XlsxSharp.IO.CodeGen.Model.Elements;

namespace XlsxSharp.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:group/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:group name="EG_ExtensionList" >
///   <xsd:sequence>
///     <xsd:element name = "ext" type="CT_Extension" minOccurs="0" maxOccurs="unbounded"/>
///   </xsd:sequence>
/// </xsd:group>
/// ]]></code>
/// </example>
/// </summary>
public class GroupDefinition : IParslet
{
    public required ParsletName Name { get; init; }

    public required IElementGroup Content { get; init; }

    void IParslet.GenerateParseMethod(CodeBuilder code)
    {
        if (this.Content is Choice choice)
        {
            ElementsCount choicesCount = choice.DetermineChoicesCount();
            if (choicesCount != ElementsCount.OneToOne)
            {
                throw new NotSupportedException("Element group choice should have 1 occurence.");
            }

            string? returnCsType = code.StartParseMethod(this.Name);
            code.OpenBrace();
            List<Variable> variables = choice.GenerateParseContent(
                this.Name,
                choicesCount,
                code,
                throwOnFail: false
            );
            if (returnCsType is null)
            {
                code.EndLine();
                code.WriteIndent().AppendCallHook(this.Name, variables).Append(";").EndLine();
                code.AddLine("return Xpr.Success();");
                code.CloseBrace();
                code.EndLine();
                code.AddHookSignature(this.Name, variables);
            }
            else
            {
                code.WriteIndent()
                    .Append("return Xpr.From(")
                    .AppendCallHook(this.Name, variables)
                    .Append(");")
                    .EndLine();
                code.CloseBrace();
            }
        }
        else
        {
            throw new NotImplementedException("Only choice implemented.");
        }
    }
}
