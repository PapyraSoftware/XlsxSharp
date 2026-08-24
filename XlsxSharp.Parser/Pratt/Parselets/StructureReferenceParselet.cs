namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// A bare structure reference (<see cref="TokenType.SquareIdent"/>) with no table name, e.g.
/// <c>[Column]</c> or <c>[#Totals]</c> - only valid for a formula entered directly in the table
/// (e.g. a totals row). A table-qualified structure reference (<c>Table1[Column]</c>) is handled
/// by <see cref="IdentParselet{TScalar,T,TContext}"/> instead, since it starts with a name.
/// </summary>
/// <remarks>
/// Delegates the actual bracket-content parsing to <see cref="TokenParser.ParseIntraTableReference"/>,
/// the same routine <see cref="FormulaParser{TScalarValue,TNode,TContext}"/> uses - the pratt
/// lexer's <see cref="TokenType.SquareIdent"/> token covers exactly the same bracket syntax
/// (including up to two levels of nested brackets), so there's no reason to duplicate that logic.
/// </remarks>
internal class StructureReferenceParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public StructureReferenceParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        TokenParser.ParseIntraTableReference(
            token.GetText(this._parser.Input),
            out StructuredReferenceArea area,
            out string? firstColumn,
            out string? lastColumn
        );
        T value = this._factory.StructureReference(ctx, token.Range, area, firstColumn, lastColumn ?? firstColumn);
        return new Node<T>(value, token.Range);
    }
}
