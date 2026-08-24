namespace XlsxSharp.Parser.Ast;

public record BangReferenceNode(ReferenceArea Reference) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"!{Reference.GetDisplayString(style)}";
    }
}
