namespace XlsxSharp.Parser.Ast;

public record ExternalSheetReferenceNode(int WorkbookIndex, string Sheet, ReferenceArea Reference)
    : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"[{WorkbookIndex}]{Sheet}!{Reference.GetDisplayString(style)}";
    }
}
