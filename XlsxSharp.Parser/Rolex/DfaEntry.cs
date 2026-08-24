namespace XlsxSharp.Parser.Rolex;

internal struct DfaEntry
{
    public DfaTransitionEntry[] Transitions;
    public int AcceptSymbolId;
    public DfaEntry(DfaTransitionEntry[] transitions, int acceptSymbolId)
    {
        this.Transitions = transitions;
        this.AcceptSymbolId = acceptSymbolId;
    }
}