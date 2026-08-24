using System.Diagnostics.CodeAnalysis;
using System.Text;
using XlsxSharp.IO.CodeGen.Model;

namespace XlsxSharp.IO.CodeGen;

internal class CodeBuilder
{
    /// <summary>
    /// C# keywords. The variables with that name must be escaped, e.g. <c>in</c> must be
    /// <c>@in</c>.
    /// </summary>
    private static readonly HashSet<string> Keywords = ["in", "out", "ref", "char"];

    private readonly SchemeTypeMap _typeMap;
    private readonly StringBuilder _sb;
    private int _indentLevel;

    public CodeBuilder(StringBuilder sb, SchemeTypeMap typeMap)
    {
        this._sb = sb;
        this._typeMap = typeMap;
    }

    internal CodeBuilder AddLine(string lineText)
    {
        this.AddIndentedLine(lineText);
        return this;
    }

    internal CodeBuilder OpenBrace()
    {
        this.AddIndentedLine("{");
        this._indentLevel++;
        return this;
    }

    internal CodeBuilder CloseBrace()
    {
        this._indentLevel--;
        this.AddIndentedLine("}");
        return this;
    }

    internal CodeBuilder Append(string text)
    {
        this._sb.Append(text);
        return this;
    }

    internal CodeBuilder EndLine()
    {
        this._sb.AppendLine();
        return this;
    }

    internal CodeBuilder WriteIndent()
    {
        this.AddIndentation();
        return this;
    }

    internal CodeBuilder AppendVariable(string variableName)
    {
        this._sb.Append(Keywords.Contains(variableName) ? '@' + variableName : variableName);
        return this;
    }

    internal string? StartParseMethod(ParsletName name, params string[] parameters)
    {
        string parseMethodReturnType;
        if (!this.TryGetCsType(name, out string? csReturnType))
        {
            csReturnType = null;
            parseMethodReturnType = "Xpr";
        }
        else
        {
            parseMethodReturnType = "Xpr<" + csReturnType + ">";
        }

        this.AddIndentedLine(
            $"private {parseMethodReturnType} Parse{name.WithoutPrefix()}({string.Join(", ", parameters)})"
        );
        return csReturnType;
    }

    internal string GetSimpleType(string simpleType) =>
        this._typeMap.GetSimpleType(simpleType).CsTypeName;

    internal CodeBuilder AppendValue(string simpleType, string value)
    {
        string mappedValue = this._typeMap.GetSimpleType(simpleType).MapValue(value);
        this._sb.Append(mappedValue);
        return this;
    }

    internal bool TryGetCsType(ParsletName name, [NotNullWhen(true)] out string? csType) =>
        this._typeMap.TryGetParsletCsType(name, out csType);

    internal string GetCsItemType(ParsletName parsletName)
    {
        if (this._typeMap.TryItemGetValue(parsletName, out string? code))
        {
            return code;
        }

        throw new KeyNotFoundException($"Missing parslet '{parsletName.Value}'");
    }

    internal CodeBuilder AppendCtParseCall(ParsletName name, string elementName) =>
        this.AppendParseCall(name, ["\"" + elementName + "\"", "_ns"]);

    internal CodeBuilder AppendGroupParseCall(ParsletName name)
    {
        string parseCall = this._typeMap.GetParseCall(name);
        return this.Append($"{parseCall}()");
    }

    internal Variable? AddGroupParseCall(ParsletName name, string storeVariableName)
    {
        if (!this.TryGetCsType(name, out string? csType))
        {
            this.WriteIndent().AppendParseCall(name, []).Append(";").EndLine();
            return null;
        }

        this.WriteIndent()
            .Append("var ")
            .AppendVariable(storeVariableName)
            .Append(" = ")
            .AppendParseCall(name, [])
            .Append(";")
            .EndLine();
        return new Variable(csType, storeVariableName);
    }

    private CodeBuilder AppendParseCall(ParsletName name, string[] arguments)
    {
        string parseCall = this._typeMap.GetParseCall(name);
        return this.Append($"{parseCall}({string.Join(", ", arguments)})");
    }

    internal CodeBuilder AppendCallHook(ParsletName name, IReadOnlyList<Variable> arguments)
    {
        this.Append("On").Append(name.WithoutPrefix()).Append("Parsed(");
        bool isFirst = true;
        foreach (Variable variable in arguments)
        {
            if (!isFirst)
            {
                this.Append(", ");
            }

            this.AppendVariable(variable.Name);
            isFirst = false;
        }

        this.Append(")");
        return this;
    }

    internal CodeBuilder AddHookSignature(ParsletName name, IReadOnlyList<Variable> parameters)
    {
        this.WriteIndent().Append("partial void On").Append(name.WithoutPrefix()).Append("Parsed(");

        bool isFirst = true;
        foreach (Variable parameter in parameters)
        {
            if (!isFirst)
            {
                this.Append(", ");
            }

            this.Append(parameter.Type).Append(" ").AppendVariable(parameter.Name);
            isFirst = false;
        }

        this.Append(");").EndLine();
        return this;
    }

    internal CodeBuilder AppendSimpleTypeMethod(AttributeElement attribute)
    {
        string codeFragment = this._typeMap.GetSimpleTypeMethod(attribute);
        return this.Append(codeFragment);
    }

    private void AddIndentedLine(string text)
    {
        this.AddIndentation();
        this._sb.AppendLine(text);
    }

    private void AddIndentation()
    {
        for (int i = 0; i < this._indentLevel; i++)
        {
            this._sb.Append("    ");
        }
    }

    public override string ToString() => this._sb.ToString();
}
