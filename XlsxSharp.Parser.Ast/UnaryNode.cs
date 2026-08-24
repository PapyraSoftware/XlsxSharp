namespace XlsxSharp.Parser.Ast;

public record UnaryNode(UnaryOperation Operation) : AstNode
{
    public UnaryNode(UnaryOperation operation, AstNode child)
        : this(operation)
    {
        this.Children = [child];
    }

    public override string GetDisplayString(ReferenceStyle style)
    {
        return Operation switch
        {
            UnaryOperation.Percent => "%",
            UnaryOperation.Minus => "-",
            UnaryOperation.Plus => "+",
            UnaryOperation.ImplicitIntersection => "Implicit intersection",
            UnaryOperation.SpillRange => "Spill",
            _ => throw new NotSupportedException(),
        };
    }
};
