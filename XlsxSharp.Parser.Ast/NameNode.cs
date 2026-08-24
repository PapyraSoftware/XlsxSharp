namespace XlsxSharp.Parser.Ast;

public record NameNode(string Name) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return Name;
    }
}
