using System.Text;
using XlsxSharp.IO.CodeGen.Model;
using XlsxSharp.IO.CodeGen.Model.TopLevel;

namespace XlsxSharp.IO.CodeGen;

public class ParserGenerator
{
    private readonly Schema _schema;
    private readonly string _readerName;
    private readonly List<ParsletName> _parseMethods = [];
    private readonly SchemeTypeMap _typeMap;
    private readonly List<string> _usings = [];
    private string _targetNamespace = "XlsxSharp.Excel.IO";

    public ParserGenerator(Schema schema, SchemeTypeMap typeMap, string readerField)
    {
        this._schema = schema;
        this._typeMap = typeMap;
        this._readerName = readerField;
    }

    public ParserGenerator WithNamespace(string targetNamespace)
    {
        this._targetNamespace = targetNamespace;
        return this;
    }

    public ParserGenerator AddUsing(string usingNamespace)
    {
        this._usings.Add(usingNamespace);
        return this;
    }

    /// <summary>
    /// Generate <c>Parse*</c> method for a top-level element in the XSD file.
    /// </summary>
    /// <param name="name">Name of a complex type or element group.</param>
    public ParserGenerator AddParseMethod(ParsletName name)
    {
        if (this._parseMethods.Contains(name))
        {
            throw new InvalidOperationException($"Parse method for {name} was already added.");
        }

        this._parseMethods.Add(name);
        return this;
    }

    /// <summary>
    /// Generate code from the configuration and a XML schema.
    /// </summary>
    /// <returns>Generated source code.</returns>
    public string Generate()
    {
        CodeBuilder code = new(new StringBuilder(), this._typeMap);
        code.AddLine("#nullable enable");
        code.EndLine();
        foreach (string usingNs in this._usings)
        {
            code.AddLine($"using {usingNs};");
        }

        code.EndLine();
        code.AddLine($"namespace {this._targetNamespace};");
        code.EndLine();
        code.AddLine($"internal partial class {this._readerName}");
        code.OpenBrace();

        if (this._parseMethods.Count > 0)
        {
            this.GenerateParseMethod(code, this._parseMethods[0]);
        }

        foreach (ParsletName parseMethod in this._parseMethods[1..])
        {
            code.EndLine();
            this.GenerateParseMethod(code, parseMethod);
        }

        code.CloseBrace();
        return code.ToString();
    }

    private void GenerateParseMethod(CodeBuilder code, ParsletName parsletName)
    {
        if (!this._schema.TryGetParslet(parsletName, out IParslet? parslet))
        {
            throw new InvalidOperationException(
                $"Unable to find definition for '{parsletName.Value}'. Was it part of the XSD file?"
            );
        }

        parslet.GenerateParseMethod(code);
    }
}
