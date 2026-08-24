namespace XlsxSharp.Parser.Ast;

public record CellFunctionNode(RowCol RowCol) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style) => RowCol.GetDisplayString(style);
}
