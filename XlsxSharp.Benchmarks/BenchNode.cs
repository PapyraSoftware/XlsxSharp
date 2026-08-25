namespace XlsxSharp.Benchmarks;

/// <summary>
/// The lightest possible AST node: one allocation tagged with the grammar rule that produced it.
/// Both <see cref="XlsxSharpAstFactory"/> and <see cref="ClosedXmlAstFactory"/> build this same
/// shape, so <see cref="FormulaParserBenchmarks"/> measures the parsers themselves rather than
/// differences in how much a fancier factory would allocate per node.
/// </summary>
internal sealed class BenchNode(string kind)
{
    public string Kind { get; } = kind;
}
