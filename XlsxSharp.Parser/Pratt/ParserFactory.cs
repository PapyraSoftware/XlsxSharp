using XlsxSharp.Parser.Pratt.Parselets;

namespace XlsxSharp.Parser.Pratt;

internal static class ParserFactory
{
    public static Parser<TNode, TContext> Create<TScalar, TNode, TContext>(
        IAstFactory<TScalar, TNode, TContext> factory)
    {
        Parser<TNode, TContext> parser = new();

        // Register prefix parselets
        parser.Register(TokenType.Number, new NumberParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.LeftParen, new GroupParselet<TNode, TContext>(parser));
        parser.Register(TokenType.Ident, new IdentParselet<TScalar,TNode,TContext>(factory, parser));
        parser.Register(TokenType.QIdent, new QIdentParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Text, new TextParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Error, new ErrorParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.SquareIdent, new StructureReferenceParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.Plus, new UnaryOpParselet<TScalar, TNode, TContext>(factory, parser, UnaryOperation.Plus));
        parser.Register(TokenType.Minus, new UnaryOpParselet<TScalar, TNode, TContext>(factory, parser, UnaryOperation.Minus));

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
