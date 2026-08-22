using System;
using System.Collections.Generic;
using XlsxSharp.IO.CodeGen.Model.TopLevel;

namespace XlsxSharp.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:group ref="some:group">]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// </summary>
public class GroupReference : ILeafElement
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// A reference to the element (<see cref="GroupDefinition.Name"/>).
    /// </summary>
    public required string RefName { get; init; }

    public required Occurrences Occurrences { get; init; }

    internal List<Variable> GenerateSequenceParseCall(CodeBuilder code)
    {
        if (this.Occurrences.HasFixedCount)
        {
            List<Variable> variables = [];
            for (int i = 0; i < this.Occurrences.ActualMax; ++i)
            {
                // GroupReference is only called from a sequence. Therefore we do know that
                // we are in the opened sequence element and that there must be an element
                // of cardinality 1-1. Therefore, if it isn't found, it's unrecoverable error.
                if (!code.TryGetCsType(this.RefName, out string? csGroupType))
                {
                    code.WriteIndent()
                        .Append("if (")
                        .AppendGroupParseCall(this.RefName)
                        .Append(" is { IsFail: true })")
                        .EndLine();
                    code.OpenBrace();
                    code.AddLine(
                        $"throw PartStructureException.ExpectedElementNotFound(\"{this.RefName}\", _reader);"
                    );
                    code.CloseBrace();
                }
                else
                {
                    // If there is only 1 item, don't add suffix for the variable
                    string variableSuffix =
                        this.Occurrences.ActualMax == 1 ? string.Empty : i.ToString();

                    string variableName =
                        char.ToLowerInvariant(this.RefName[3]) + this.RefName[4..];
                    string itemVariableName = variableName + variableSuffix;
                    string resultVariable = variableName + "Result" + variableSuffix;

                    code.AddLine($"{csGroupType} {itemVariableName};");
                    code.WriteIndent()
                        .Append("if (")
                        .AppendGroupParseCall(this.RefName)
                        .Append(" is { IsSuccess: true } ")
                        .Append(resultVariable)
                        .Append(")")
                        .EndLine();
                    code.OpenBrace();
                    code.AddLine($"{itemVariableName} = {resultVariable}.Value;");
                    code.CloseBrace();
                    code.AddLine("else");
                    code.OpenBrace();
                    code.AddLine("throw PartStructureException.ExpectedElementNotFound();");
                    code.CloseBrace();
                    variables.Add(new Variable(csGroupType, itemVariableName));
                }
            }

            return variables;
        }

        switch (this.Occurrences.Elements)
        {
            case ElementsCount.ZeroToOne:
            {
                if (!code.TryGetCsType(this.RefName, out string? csGroupType))
                {
                    // if (ParseSomeGroup() is { IsSuccess: true })
                    // {
                    //     // Successfully parsed optional group reference 'EG_SomeGroup'
                    // }
                    code.WriteIndent()
                        .Append("if (")
                        .AppendGroupParseCall(this.RefName)
                        .Append(" is { IsSuccess: true })")
                        .EndLine();
                    code.OpenBrace();
                    code.AddLine(
                        $"// Successfully parsed optional group reference '{this.RefName}'"
                    );
                    code.CloseBrace();
                    return [];
                }
                else
                {
                    // GroupResultType? someGroup = default;
                    // if (ParseSomeGroup() is { IsSuccess: true } someGroupXpr)
                    // {
                    //     someGroup = someGroupXpr.Value;
                    // }
                    Variable itemVariable = new(
                        csGroupType + '?',
                        char.ToLowerInvariant(this.RefName[3]) + this.RefName[4..]
                    );
                    code.AddLine($"{csGroupType}? {itemVariable.Name} = default;");
                    code.WriteIndent()
                        .Append("if (")
                        .AppendGroupParseCall(this.RefName)
                        .Append(" is { IsSuccess: true }")
                        .Append($" {itemVariable.Name}Xpr)")
                        .EndLine();
                    code.OpenBrace();
                    code.AddLine($"{itemVariable.Name} = {itemVariable.Name}Xpr.Value;");
                    code.CloseBrace();
                    return [itemVariable];
                }
            }
            case ElementsCount.ZeroToMany:
            {
                if (code.TryGetCsType(this.RefName, out string? csGroupType))
                {
                    // Declare list variable to hold all each parsed result of EG
                    // var list = new List<ItemRes>();
                    // var item = ParseGroup();
                    // while (item.IsSuccess)
                    // {
                    //     list.Add(item.Value);
                    //     item = ParseGroup();
                    // }
                    string listVariableName =
                        char.ToLowerInvariant(this.RefName[3]) + this.RefName[4..] + "List";
                    string listVariableType = $"List<{csGroupType}>";
                    code.WriteIndent()
                        .Append("var ")
                        .AppendVariable(listVariableName)
                        .Append(" = ")
                        .Append($"new {listVariableType}()")
                        .Append(";")
                        .EndLine();

                    string itemVariableName =
                        char.ToLowerInvariant(this.RefName[3]) + this.RefName[4..];
                    code.AddGroupParseCall(this.RefName, itemVariableName);
                    code.WriteIndent()
                        .Append("while (")
                        .AppendVariable(itemVariableName)
                        .Append(".IsSuccess)")
                        .EndLine();
                    code.OpenBrace();
                    code.WriteIndent()
                        .AppendVariable(listVariableName)
                        .Append(".Add(")
                        .AppendVariable(itemVariableName)
                        .Append(".Value)")
                        .Append(";")
                        .EndLine();
                    code.WriteIndent()
                        .AppendVariable(itemVariableName)
                        .Append(" = ")
                        .AppendGroupParseCall(this.RefName)
                        .Append(";")
                        .EndLine();
                    code.CloseBrace();

                    return [new Variable(listVariableType, listVariableName)];
                }
                else
                {
                    // Declare list variable to hold all each parsed result of EG
                    // while (ParseGroup() is { IsSuccess: true })
                    // {
                    //     // Parsed 'EG_SomeGroup'
                    // }
                    code.WriteIndent()
                        .Append("while (")
                        .AppendGroupParseCall(this.RefName)
                        .Append(" is { IsSuccess: true })")
                        .EndLine();
                    code.OpenBrace();
                    code.AddLine($"// Parsed group element {this.RefName} with cardinality 0..N");
                    code.CloseBrace();
                    return [];
                }
            }
            default:
                throw new NotImplementedException();
        }
    }
}
