namespace XlsxSharp.IO;

/// <summary>
/// An info sent to the consuming application in case of a mismatch.
/// </summary>
public class MismatchInfo
{
    /// <summary>
    /// Position in a document where mismatch happened.
    /// </summary>
    public LineInfo LineInfo { get; init; }
}
