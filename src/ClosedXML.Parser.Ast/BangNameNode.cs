namespace ClosedXML.Parser;

public record BangNameNode(string Name) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"!{Name}";
    }
}