namespace XlsxSharp.Parser.Rolex;

internal struct DfaTransitionEntry
{
    public int[] PackedRanges;
    public int Destination;

    public DfaTransitionEntry(int[] packedRanges, int destination)
    {
        this.PackedRanges = packedRanges;
        this.Destination = destination;
    }
}