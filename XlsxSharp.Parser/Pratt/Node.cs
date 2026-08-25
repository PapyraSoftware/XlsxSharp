namespace XlsxSharp.Parser.Pratt;

/// <summary>
/// An info about node used during parsing.
/// </summary>
/// <typeparam name="T">The <c>TNode</c> type of a node from <see cref="IAstFactory{TScalarValue,TNode,TContext}"/>.</typeparam>
internal readonly struct Node<T>
{
    public Node(T value, int start, int end)
        : this(value, new SymbolRange(start, end))
    {
    }

    public Node(T value, SymbolRange range)
        : this(value, range, isPureReference: false)
    {
    }

    public Node(T value, SymbolRange range, bool isPureReference)
    {
        this.Value = value;
        this.Range = range;
        this.IsPureReference = isPureReference;
    }

    /// <summary>
    /// Parsed value of a node, created by the <see cref="IAstFactory{TScalarValue,TNode,TContext}"/>.
    /// </summary>
    public T Value { get; }

    /// <summary>
    /// A range that was used to created the node.
    /// </summary>
    public SymbolRange Range { get; }

    /// <summary>
    /// Is this node one that, standing alone, is reference-shaped (a cell/area reference, a
    /// defined name, a structure reference, a 3D/external reference, an unresolved <c>#REF!</c>,
    /// or a call to one of the five functions the oracle's lexer recognizes as capable of
    /// returning a reference - <c>CHOOSE</c>, <c>IF</c>, <c>INDEX</c>, <c>INDIRECT</c>,
    /// <c>OFFSET</c>)? Only such nodes can be an operand of the range operator (<c>:</c>) - see
    /// <see cref="Parser{T,TContext}.ParseRangeChain"/>. Any operator applied on top (unary,
    /// percent, power, ...) makes a node no longer pure - it's tracked per-node instead of
    /// per-token-type because e.g. a parenthesized expression is only pure when its content is.
    /// </summary>
    public bool IsPureReference { get; }

    public static implicit operator T(Node<T> node)
    {
        return node.Value;
    }
}
