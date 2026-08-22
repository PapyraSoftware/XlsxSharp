using System.Collections.Generic;
using XlsxSharp.IO.CodeGen.Model.Elements;

namespace XlsxSharp.IO.CodeGen.Model.TopLevel;

/// <summary>
/// Base class for nodes representing a <c><![CDATA[<xsd:compleType>]]></c>.
/// </summary>
public abstract class ComplexType : IParslet
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required ParsletName Name { get; set; }

    public List<OneOf<AttributeElement, AttributeGroupReference>> Attributes { get; set; } = [];

    /// <summary>
    /// Can text be freely interspersed with elements? Only used when <c>complexType</c> contains
    /// <c>any</c>.
    /// </summary>
    public required bool? Mixed { get; init; }

    void IParslet.GenerateParseMethod(CodeBuilder code)
    {
        List<Variable> attributeVariables = [];
        string? csReturnType = code.StartParseMethod(this.Name, "string elementName", "string ns");
        code.OpenBrace();

        // If we are not in correct element, don't continue onwards.
        code.WriteIndent().Append("if (!_reader.TryOpen(elementName, ns))").EndLine();
        code.OpenBrace();
        if (csReturnType is not null)
        {
            code.AddLine($"return Xpr.Fail<{csReturnType}>();");
        }
        else
        {
            code.AddLine("return Xpr.Fail();");
        }
        code.CloseBrace();
        code.EndLine();

        foreach (OneOf<AttributeElement, AttributeGroupReference> oneOfAttribute in this.Attributes)
        {
            if (
                oneOfAttribute.TryPickT1(
                    out AttributeElement? attribute,
                    out AttributeGroupReference? attributeGroup
                )
            )
            {
                Variable attributeVariable = attribute.Generate(code);
                attributeVariables.Add(attributeVariable);
            }
            else
            {
                // There are only 2 attribute groups, just hand-code it manually in the reader.
                string agParseMethod = "Parse" + attributeGroup.RefName[3..];
                if (code.TryGetCsType(attributeGroup.RefName, out string? csType))
                {
                    string agVarName =
                        char.ToLowerInvariant(attributeGroup.RefName[3])
                        + attributeGroup.RefName[4..];
                    code.AddLine($"{csType} {agVarName} = {agParseMethod}();");
                    attributeVariables.Add(new Variable(csType, agVarName));
                }
                else
                {
                    code.AddLine($"{agParseMethod}();");
                }
            }
        }

        if (this.Attributes.Count > 0)
        {
            code.EndLine();
        }

        List<Variable> elementVariables = this.GenerateParseMethod(code);
        List<Variable> dataVariables = [.. elementVariables, .. attributeVariables];

        code.AddLine("_reader.Close(elementName, ns);");
        code.EndLine();

        if (csReturnType is null)
        {
            code.WriteIndent().AppendCallHook(this.Name, dataVariables).Append(";").EndLine();
            code.AddLine("return Xpr.Success();");
            code.CloseBrace();
            code.EndLine();
            code.AddHookSignature(this.Name, dataVariables);
        }
        else
        {
            // If the Parse* method should map to a value, it's not possible to use partial hook.
            // Partial methods can't return value. The method will be displayed as uncompilable,
            // which is desirable, so it is implemented in the partial reader class by the developer.
            code.WriteIndent()
                .Append("return Xpr.From(")
                .AppendCallHook(this.Name, dataVariables)
                .Append(");")
                .EndLine();
            code.CloseBrace();
        }
    }

    internal abstract List<Variable> GenerateParseMethod(CodeBuilder code);
}
