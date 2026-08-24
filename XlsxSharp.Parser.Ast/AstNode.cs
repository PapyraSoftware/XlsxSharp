namespace XlsxSharp.Parser.Ast;

public abstract record AstNode
{
    public AstNode[] Children { get; init; } = Array.Empty<AstNode>();

    /// <summary>
    /// Render node and its children in a reference style.
    /// </summary>
    public abstract string GetDisplayString(ReferenceStyle style);

    public virtual string GetTypeString() => this.GetType().Name[..^4]; // Strip Node suffix

    public virtual bool Equals(AstNode? other) =>
        other is not null && this.Children.SequenceEqual(other.Children);

    public override int GetHashCode()
    {
        int hash = 0;
        foreach (AstNode child in this.Children)
        {
            unchecked
            {
                hash += child.GetHashCode();
            }
        }
        return hash;
    }
}
