namespace XlsxSharp.Parser.Tests;

public class AstFactoryTests
{
    [Test]
    [Arguments("1+TRUE", 2, 6)]
    public async Task LogicalRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new LogicalVisitor());
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1:#REF!", 3, 8)]
    public async Task ErrorRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ErrorVisitor());
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("JOIN(1.15)", 5, 9)]
    public async Task NumberRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new NumberVisitor());
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("JOIN(\"A b c\")", 5, 12)]
    public async Task TextRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new TextVisitor());
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("SUM( {1,2,3})", 5, 12)]
    public async Task ArrayRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ArrayVisitor());
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1", 0, 2)]
    [Arguments("A1 ", 0, 2)]
    [Arguments(" A1 ", 1, 3)]
    [Arguments(" $B7:D$18 ", 1, 9)]
    [Arguments("  1:7", 2, 5)]
    [Arguments("SUM(A:C)", 4, 7)]
    public async Task ReferenceRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ReferenceVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("Sheet!A1", 0, 8)]
    [Arguments(" S!$A$1:$B4 ", 1, 11)]
    [Arguments("1+'Johnny''s'!Z26", 2, 17)]
    public async Task SheetReferenceRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new SheetReferenceVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("name + !$A$1:$B4 ", 7, 16)]
    [Arguments("1+!$Z$26", 2, 8)]
    public async Task BangReferenceRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new BangReferenceVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("Jan:Feb!A1", 0, 10)]
    [Arguments("1+Zara:Beta!$A$1:$B4+4", 2, 20)]
    [Arguments("1+'2022 Q1:2024 Q1'!Z26", 2, 23)]
    public async Task Reference3DRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new Reference3DVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("1+[2]S!A1 + 2", 2, 9)]
    [Arguments("1+'[2]D and D'!A1*2", 2, 17)]
    public async Task ExternalSheetReferenceRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalSheetReferenceVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("1+[2]A:Z!A1 + 2", 2, 11)]
    [Arguments("1+'[2]D and D:B and B'!A1*2", 2, 25)]
    public async Task ExternalReference3DRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalReference3DVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("2+A1(14,7)+2", 2, 10)]
    public async Task CellFunctionRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new CellFunctionVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("2+[Column]", 2, 10)]
    public async Task StructureReferenceNoTableRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new StructureReferenceNoTableVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("2+Table1[Column]+3", 2, 16)]
    public async Task StructureReferenceRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new StructureReferenceVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("2+[1]!Table1[Column]+3", 2, 20)]
    public async Task ExternalStructureReferenceRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalStructureReferenceVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1+SUM(4)+name", 3, 9)]
    public async Task FunctionRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new FunctionVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1+[1]!SUM(4)+name", 3, 13)]
    public async Task ExternalFunctionRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalFunctionVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1+Sheet!SUM(4)+name", 3, 15)]
    public async Task SheetFunctionRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new SheetFunctionVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1+[5]Sheet!SUM(4)+name", 3, 18)]
    public async Task ExternalSheetFunctionRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalSheetFunctionVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("A1+TRUE-name", 8, 12)]
    public async Task NameRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new NameVisitor());
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("Sheet!name + 4", 0, 10)]
    public async Task SheetNameRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new SheetNameVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("+[7]!name + 4", 1, 9)]
    public async Task ExternalNameRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalNameVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    [Arguments("+[7]Sheet!name + 4", 1, 14)]
    public async Task ExternalSheetNameRange(string formula, int start, int end)
    {
        Result result = new();
        FormulaParser<object?, string, Result>.CellFormulaA1(
            formula,
            result,
            new ExternalSheetNameVisitor()
        );
        await Assert.That(result.Value).IsEqualTo(new SymbolRange(start, end));
    }

    [Test]
    public async Task BinaryOperationRange()
    {
        List<SymbolRange> result = new();
        FormulaParser<object?, string, List<SymbolRange>>.CellFormulaA1(
            "1+2+3+4",
            result,
            new BinaryOperationVisitor()
        );
        await Assert
            .That(result)
            .IsEquivalentTo(
                new[] { new SymbolRange(0, 3), new SymbolRange(0, 5), new SymbolRange(0, 7) }
            );
    }

    [Test]
    public async Task UnaryOperationRange()
    {
        List<SymbolRange> result = new();
        FormulaParser<object?, string, List<SymbolRange>>.CellFormulaA1(
            "-7+8%",
            result,
            new UnaryOperationVisitor()
        );
        await Assert
            .That(result)
            .IsEquivalentTo(new[] { new SymbolRange(0, 2), new SymbolRange(3, 5) });
    }

    [Test]
    public async Task NestedRange()
    {
        List<SymbolRange> result = new();
        FormulaParser<object?, string, List<SymbolRange>>.CellFormulaA1(
            "-( 1 + (2))+8%",
            result,
            new NestedOperationVisitor()
        );
        await Assert
            .That(result)
            .IsEquivalentTo(new[] { new SymbolRange(7, 10), new SymbolRange(1, 11) });
    }

    private class LogicalVisitor : BaseVisitor
    {
        public override string LogicalNode(Result context, SymbolRange range, bool logical)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ErrorVisitor : BaseVisitor
    {
        public override string ErrorNode(
            Result context,
            SymbolRange range,
            ReadOnlySpan<char> error
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class NumberVisitor : BaseVisitor
    {
        public override string NumberNode(Result context, SymbolRange range, double number)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class TextVisitor : BaseVisitor
    {
        public override string TextNode(Result context, SymbolRange range, string text)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ArrayVisitor : BaseVisitor
    {
        public override string ArrayNode(
            Result context,
            SymbolRange range,
            int rows,
            int columns,
            IReadOnlyList<object?> elements
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ReferenceVisitor : BaseVisitor
    {
        public override string Reference(Result context, SymbolRange range, ReferenceArea reference)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class SheetReferenceVisitor : BaseVisitor
    {
        public override string SheetReference(
            Result context,
            SymbolRange range,
            string sheet,
            ReferenceArea reference
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class BangReferenceVisitor : BaseVisitor
    {
        public override string BangReference(
            Result context,
            SymbolRange range,
            ReferenceArea reference
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class Reference3DVisitor : BaseVisitor
    {
        public override string Reference3D(
            Result context,
            SymbolRange range,
            string firstSheet,
            string lastSheet,
            ReferenceArea reference
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalSheetReferenceVisitor : BaseVisitor
    {
        public override string ExternalSheetReference(
            Result context,
            SymbolRange range,
            int workbookIndex,
            string sheet,
            ReferenceArea reference
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalReference3DVisitor : BaseVisitor
    {
        public override string ExternalReference3D(
            Result context,
            SymbolRange range,
            int workbookIndex,
            string firstSheet,
            string lastSheet,
            ReferenceArea reference
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class CellFunctionVisitor : BaseVisitor
    {
        public override string CellFunction(
            Result context,
            SymbolRange range,
            RowCol cell,
            IReadOnlyList<string> arguments
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class StructureReferenceNoTableVisitor : BaseVisitor
    {
        public override string StructureReference(
            Result context,
            SymbolRange range,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class StructureReferenceVisitor : BaseVisitor
    {
        public override string StructureReference(
            Result context,
            SymbolRange range,
            string table,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalStructureReferenceVisitor : BaseVisitor
    {
        public override string ExternalStructureReference(
            Result context,
            SymbolRange range,
            int workbookIndex,
            string table,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class FunctionVisitor : BaseVisitor
    {
        public override string Function(
            Result context,
            SymbolRange range,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<string> arguments
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalFunctionVisitor : BaseVisitor
    {
        public override string ExternalFunction(
            Result context,
            SymbolRange range,
            int workbookIndex,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<string> arguments
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class SheetFunctionVisitor : BaseVisitor
    {
        public override string Function(
            Result context,
            SymbolRange range,
            string sheetName,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<string> arguments
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalSheetFunctionVisitor : BaseVisitor
    {
        public override string ExternalFunction(
            Result context,
            SymbolRange range,
            int workbookIndex,
            string sheetName,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<string> arguments
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class NameVisitor : BaseVisitor
    {
        public override string Name(Result context, SymbolRange range, string name)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class SheetNameVisitor : BaseVisitor
    {
        public override string SheetName(
            Result context,
            SymbolRange range,
            string sheetName,
            string name
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalNameVisitor : BaseVisitor
    {
        public override string ExternalName(
            Result context,
            SymbolRange range,
            int workbookIndex,
            string name
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalSheetNameVisitor : BaseVisitor
    {
        public override string ExternalSheetName(
            Result context,
            SymbolRange range,
            int workbookIndex,
            string sheet,
            string name
        )
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class BinaryOperationVisitor : BaseVisitor<object?, string, List<SymbolRange>>
    {
        public BinaryOperationVisitor()
            : base(null, string.Empty) { }

        public override string BinaryNode(
            List<SymbolRange> context,
            SymbolRange range,
            BinaryOperation operation,
            string leftNode,
            string rightNode
        )
        {
            context.Add(range);
            return string.Empty;
        }
    }

    private class UnaryOperationVisitor : BaseVisitor<object?, string, List<SymbolRange>>
    {
        public UnaryOperationVisitor()
            : base(null, string.Empty) { }

        public override string Unary(
            List<SymbolRange> context,
            SymbolRange range,
            UnaryOperation operation,
            string node
        )
        {
            context.Add(range);
            return string.Empty;
        }
    }

    private class NestedOperationVisitor : BaseVisitor<object?, string, List<SymbolRange>>
    {
        public NestedOperationVisitor()
            : base(null, string.Empty) { }

        public override string Nested(List<SymbolRange> context, SymbolRange range, string node)
        {
            context.Add(range);
            return string.Empty;
        }
    }

    private class BaseVisitor : BaseVisitor<object?, string, Result>
    {
        protected BaseVisitor()
            : base(null, string.Empty) { }
    }

    private class BaseVisitor<TScalarValue, TNode, TContext>
        : IAstFactory<TScalarValue, TNode, TContext>
        where TNode : class
    {
        private readonly TScalarValue _defaultScalar;
        private readonly TNode _defaultNode;

        protected BaseVisitor(TScalarValue defaultScalar, TNode defaultNode)
        {
            this._defaultScalar = defaultScalar;
            this._defaultNode = defaultNode;
        }

        public virtual TScalarValue LogicalValue(TContext context, SymbolRange range, bool value)
        {
            return this._defaultScalar;
        }

        public virtual TScalarValue NumberValue(TContext context, SymbolRange range, double value)
        {
            return this._defaultScalar;
        }

        public virtual TScalarValue TextValue(TContext context, SymbolRange range, string text)
        {
            return this._defaultScalar;
        }

        public virtual TScalarValue ErrorValue(
            TContext context,
            SymbolRange range,
            ReadOnlySpan<char> error
        )
        {
            return this._defaultScalar;
        }

        public virtual TNode ArrayNode(
            TContext context,
            SymbolRange range,
            int rows,
            int columns,
            IReadOnlyList<TScalarValue> elements
        )
        {
            return this._defaultNode;
        }

        public virtual TNode BlankNode(TContext context, SymbolRange range)
        {
            return this._defaultNode;
        }

        public virtual TNode LogicalNode(TContext context, SymbolRange range, bool value)
        {
            return this._defaultNode;
        }

        public virtual TNode ErrorNode(
            TContext context,
            SymbolRange range,
            ReadOnlySpan<char> error
        )
        {
            return this._defaultNode;
        }

        public virtual TNode NumberNode(TContext context, SymbolRange range, double value)
        {
            return this._defaultNode;
        }

        public virtual TNode TextNode(TContext context, SymbolRange range, string text)
        {
            return this._defaultNode;
        }

        public virtual TNode Reference(TContext context, SymbolRange range, ReferenceArea reference)
        {
            return this._defaultNode;
        }

        public virtual TNode SheetReference(
            TContext context,
            SymbolRange range,
            string sheet,
            ReferenceArea reference
        )
        {
            return this._defaultNode;
        }

        public virtual TNode BangReference(
            TContext context,
            SymbolRange range,
            ReferenceArea reference
        )
        {
            return this._defaultNode;
        }

        public virtual TNode Reference3D(
            TContext context,
            SymbolRange range,
            string firstSheet,
            string lastSheet,
            ReferenceArea reference
        )
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalSheetReference(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            string sheet,
            ReferenceArea reference
        )
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalReference3D(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            string firstSheet,
            string lastSheet,
            ReferenceArea reference
        )
        {
            return this._defaultNode;
        }

        public virtual TNode Function(
            TContext context,
            SymbolRange range,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> arguments
        )
        {
            return this._defaultNode;
        }

        public virtual TNode Function(
            TContext context,
            SymbolRange range,
            string sheetName,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> args
        )
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalFunction(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            string sheetName,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> arguments
        )
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalFunction(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> arguments
        )
        {
            return this._defaultNode;
        }

        public virtual TNode CellFunction(
            TContext context,
            SymbolRange range,
            RowCol cell,
            IReadOnlyList<TNode> arguments
        )
        {
            return this._defaultNode;
        }

        public virtual TNode StructureReference(
            TContext context,
            SymbolRange range,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            return this._defaultNode;
        }

        public virtual TNode StructureReference(
            TContext context,
            SymbolRange range,
            string table,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalStructureReference(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            string table,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            return this._defaultNode;
        }

        public virtual TNode Name(TContext context, SymbolRange range, string name)
        {
            return this._defaultNode;
        }

        public virtual TNode SheetName(
            TContext context,
            SymbolRange range,
            string sheet,
            string name
        )
        {
            return this._defaultNode;
        }

        public TNode BangName(TContext context, SymbolRange range, string name)
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalName(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            string name
        )
        {
            return this._defaultNode;
        }

        public virtual TNode ExternalSheetName(
            TContext context,
            SymbolRange range,
            int workbookIndex,
            string sheet,
            string name
        )
        {
            return this._defaultNode;
        }

        public virtual TNode BinaryNode(
            TContext context,
            SymbolRange range,
            BinaryOperation operation,
            TNode leftNode,
            TNode rightNode
        )
        {
            return this._defaultNode;
        }

        public virtual TNode Unary(
            TContext context,
            SymbolRange range,
            UnaryOperation operation,
            TNode node
        )
        {
            return this._defaultNode;
        }

        public virtual TNode Nested(TContext context, SymbolRange range, TNode node)
        {
            return this._defaultNode;
        }
    }

    private class Result
    {
        internal SymbolRange? Value { get; set; }
    }
}
