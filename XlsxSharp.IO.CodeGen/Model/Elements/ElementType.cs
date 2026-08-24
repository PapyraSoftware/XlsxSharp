using XlsxSharp.IO.CodeGen.Model.TopLevel;

namespace XlsxSharp.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:element ref="some:element">]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>
/// (either <c><![CDATA[<xsd:sequence>]]></c> or <c><![CDATA[<xsd:choice>]]></c>).
/// <example>
/// <code><![CDATA[
///   <xsd:element name="field" maxOccurs="unbounded" type="CT_Field"/>
/// ]]></code>
/// </example>
/// </summary>
public class ElementType : IElementGroup
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// Name of the element in XML.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// A reference to a <see cref="ComplexType"/>.
    /// </summary>
    public required string TypeName { get; init; }

    public required Occurrences Occurrences { get; init; }

    internal List<Variable> GenerateSequenceParseCode(CodeBuilder code)
    {
        int min = this.Occurrences.ActualMin;
        int max = this.Occurrences.ActualMax;

        if (min == 1 && max == 1)
        {
            // If element is not found here, it's a hard fail. Don't return error, throw.
            if (code.TryGetCsType(this.TypeName, out string? csType))
            {
                Variable variable = new(csType, this.Name);
                code.WriteIndent()
                    .Append("var ")
                    .AppendVariable(variable.Name)
                    .Append(" = ")
                    .AppendCtParseCall(this.TypeName, this.Name)
                    .Append(".Value;")
                    .EndLine();
                return [variable];
            }
            else
            {
                code.WriteIndent()
                    .Append("if (")
                    .AppendCtParseCall(this.TypeName, this.Name)
                    .Append(" is { IsFail: true })")
                    .EndLine();
                code.OpenBrace();
                code.AddLine(
                    $"throw PartStructureException.ExpectedElementNotFound(\"{this.Name}\", _reader);"
                );
                code.CloseBrace();
                return [];
            }
        }
        else if (min == 0 && max == 1)
        {
            if (code.TryGetCsType(this.TypeName, out string? csType))
            {
                csType += "?";
                Variable variable = new(csType, this.Name);

                string resultVariableName = this.Name + "Result";
                code.WriteIndent()
                    .Append("var ")
                    .AppendVariable(resultVariableName)
                    .Append(" = ")
                    .AppendCtParseCall(this.TypeName, this.Name)
                    .Append(";")
                    .EndLine();

                // The default must explicitly specify a type. Ternary operator have slightly
                // unintuitive behavior for nullable types. E.g. `int? v = flag ? 1 : default`
                // would be interpreted as `int? v = flag ? 1 : 0`.
                // Interface doesn't have a default value, let's got with class. Struct doesn't really make sense.
                string defaultValue = csType.StartsWith('I') ? "null" : $"default({csType})";
                code.WriteIndent()
                    .Append("var ")
                    .AppendVariable(variable.Name)
                    .Append(" = ")
                    .AppendVariable(resultVariableName)
                    .Append(".IsSuccess ? ")
                    .AppendVariable(resultVariableName)
                    .Append($".Value : {defaultValue};")
                    .EndLine();
                return [variable];
            }
            else
            {
                code.WriteIndent()
                    .Append("if (")
                    .AppendCtParseCall(this.TypeName, this.Name)
                    .Append(" is { IsSuccess: true })")
                    .EndLine();
                code.OpenBrace();
                code.WriteIndent()
                    .Append("// Optional element '")
                    .Append(this.Name)
                    .Append("' was present")
                    .EndLine();
                code.CloseBrace();
                return [];
            }
        }

        if (min == max && min > 1 && max < int.MaxValue)
        {
            // Finite amount, but each is separate
            List<Variable> variables = [];
            for (int i = 0; i < max; i++)
            {
                if (code.TryGetCsType(this.TypeName, out string? csType))
                {
                    Variable variable = new(csType, this.Name + i);
                    code.WriteIndent()
                        .Append("var ")
                        .AppendVariable(variable.Name)
                        .Append(" = ")
                        .AppendCtParseCall(this.TypeName, this.Name)
                        .Append(".Value;")
                        .EndLine();

                    variables.Add(variable);
                }
                else
                {
                    code.WriteIndent()
                        .Append("if (")
                        .AppendCtParseCall(this.TypeName, this.Name)
                        .Append(" is { IsFail: true })")
                        .EndLine();
                    code.OpenBrace();
                    code.AddLine(
                        $"throw PartStructureException.ExpectedElementNotFound(\"{this.Name}\", _reader);"
                    );
                    code.CloseBrace();
                }
            }

            return variables;
        }

        // I am fine with few (~4) as individual elements, but more? That is just bad idea. 3 is max in Bezier and similar places
        const int threshold = 5;
        if (min < threshold && max >= threshold)
        {
            bool needsCountCheck = min > 0 || max < int.MaxValue;

            if (code.TryGetCsType(this.TypeName, out string? csType))
            {
                Variable listVariable = new($"List<{csType}>", this.Name);
                code.WriteIndent()
                    .Append("var ")
                    .AppendVariable(listVariable.Name)
                    .Append($" = new {listVariable.Type}();")
                    .EndLine();

                Variable itemVariable = new(csType, this.Name + "Item");
                code.WriteIndent()
                    .Append("while (")
                    .AppendCtParseCall(this.TypeName, this.Name)
                    .Append(" is { IsSuccess: true} ")
                    .AppendVariable(itemVariable.Name)
                    .Append(")")
                    .EndLine();
                code.OpenBrace();
                code.WriteIndent()
                    .AppendVariable(listVariable.Name)
                    .Append(".Add(")
                    .AppendVariable(itemVariable.Name)
                    .Append(".Value);")
                    .EndLine();
                code.CloseBrace();

                if (needsCountCheck)
                {
                    CountInRange(code, min, max, $"{listVariable.Name}.Count");
                }

                return [listVariable];
            }
            else
            {
                string countVariable = this.Name + "Count";
                if (needsCountCheck)
                {
                    code.AddLine($"var {countVariable} = 0;");
                }

                code.WriteIndent()
                    .Append("while (")
                    .AppendCtParseCall(this.TypeName, this.Name)
                    .Append(" is { IsSuccess: true })")
                    .EndLine();
                code.OpenBrace();
                code.AddLine(
                    $"// Parsed another element '{this.Name}' with cardinality {min}-{max}"
                );
                if (needsCountCheck)
                {
                    code.AddLine($"{countVariable}++;");
                }

                code.CloseBrace();

                if (needsCountCheck)
                {
                    CountInRange(code, min, max, countVariable);
                }

                return [];
            }
        }
        throw new NotSupportedException($"Unexpected occurence range {min}-{max}.");
    }

    private static void CountInRange(CodeBuilder code, int min, int max, string countVariable)
    {
        code.EndLine();
        if (min > 0 && max < int.MaxValue)
        {
            code.AddLine($"if ({countVariable} is < {min} or > {max})");
        }
        else if (min > 0)
        {
            code.AddLine($"if ({countVariable} < {min})");
        }
        else if (max < int.MaxValue)
        {
            code.AddLine($"if ({countVariable} > {max})");
        }
        code.OpenBrace();
        code.AddLine("throw PartStructureException.IncorrectElementsCount();");
        code.CloseBrace();
    }
}
