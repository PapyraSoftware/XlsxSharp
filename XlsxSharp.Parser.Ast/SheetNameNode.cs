namespace XlsxSharp.Parser.Ast;

public record SheetNameNode(string Sheet, string Name) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        string sheet = NameUtils.ShouldQuote(Sheet)
            ? '\'' + Sheet.Replace("'", "''") + '\''
            : Sheet;
        return $"{sheet}!{Name}";
    }
}
