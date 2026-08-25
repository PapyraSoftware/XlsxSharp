using XlsxSharp.Parser.Pratt.Parselets;

namespace XlsxSharp.Parser.Pratt;

internal static class ParserFactory
{
    public static Parser<TNode, TContext> Create<TScalar, TNode, TContext>(
        IAstFactory<TScalar, TNode, TContext> factory)
    {
        Parser<TNode, TContext> parser = new()
        {
            RangeCombiner = (ctx, range, left, right) =>
                factory.BinaryNode(ctx, range, BinaryOperation.Range, left, right),
            UnionCombiner = (ctx, range, left, right) =>
                factory.BinaryNode(ctx, range, BinaryOperation.Union, left, right),
            IntersectionCombiner = (ctx, range, left, right) =>
                factory.BinaryNode(ctx, range, BinaryOperation.Intersection, left, right),
            SpillCombiner = (ctx, range, operand) =>
                factory.Unary(ctx, range, UnaryOperation.SpillRange, operand),
            NestedCombiner = (ctx, range, operand) => factory.Nested(ctx, range, operand),
        };

        // Register prefix parselets
        IdentParselet<TScalar, TNode, TContext> identParselet = new(factory, parser);
        parser.Register(TokenType.Number, new NumberParselet<TScalar, TNode, TContext>(factory, parser, identParselet));
        parser.Register(TokenType.LeftParen, new GroupParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Ident, identParselet);
        parser.Register(TokenType.QIdent, new QIdentParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Text, new TextParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Error, new ErrorParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.SquareIdent, new StructureReferenceParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.LeftCurly, new ArrayParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Bang, new BangReferenceParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Plus, new UnaryOpParselet<TScalar, TNode, TContext>(factory, parser, UnaryOperation.Plus));
        parser.Register(TokenType.Minus, new UnaryOpParselet<TScalar, TNode, TContext>(factory, parser, UnaryOperation.Minus));
        parser.Register(TokenType.Intersection, new UnaryOpParselet<TScalar, TNode, TContext>(factory, parser, UnaryOperation.ImplicitIntersection));
        parser.Register(TokenType.Equal, new RedundantEqualsParselet<TScalar, TNode, TContext>(factory, parser));

        // Register operation parselets
        parser.Register(TokenType.Plus, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Addition, BindingPower.Addition));
        parser.Register(TokenType.Minus, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Subtraction, BindingPower.Subtraction));
        parser.Register(TokenType.Mul, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Multiplication, BindingPower.Multiplication));
        parser.Register(TokenType.Div, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Division, BindingPower.Division));
        parser.Register(TokenType.Pow, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Power, BindingPower.Exponentiation));
        parser.Register(TokenType.Concat, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Concat, BindingPower.Concat));
        parser.Register(TokenType.Equal, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Equal, BindingPower.Comparison));
        parser.Register(TokenType.NotEqual, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.NotEqual, BindingPower.Comparison));
        parser.Register(TokenType.Less, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.LessThan, BindingPower.Comparison));
        parser.Register(TokenType.LessEqual, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.LessOrEqualThan, BindingPower.Comparison));
        parser.Register(TokenType.Greater, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.GreaterThan, BindingPower.Comparison));
        parser.Register(TokenType.GreaterEqual, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.GreaterOrEqualThan, BindingPower.Comparison));
        parser.Register(TokenType.Percent, new PercentParselet<TScalar, TNode, TContext>(factory));

        return parser;
    }
}
