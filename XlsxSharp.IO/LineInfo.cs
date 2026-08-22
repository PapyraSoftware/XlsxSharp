namespace XlsxSharp.IO;

/// <summary>
/// Information about a position in a document.
/// </summary>
/// <remarks>
/// In some cases, the position in a document doesn't make sense (e.g. in-memory XML tree created
/// from objects, not from a text). The values are 0 in such a case.
/// </remarks>
/// <param name="LineNumber">Line number, can be 0 if not available.</param>
/// <param name="LinePosition">Position on the line, can be 0 if not available.</param>
public readonly record struct LineInfo(int LineNumber, int LinePosition)
{
    public override string ToString()
    {
        return $"{this.LineNumber}:{this.LinePosition}";
    }
}
