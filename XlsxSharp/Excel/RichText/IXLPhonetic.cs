namespace XlsxSharp.Excel.RichText;

public interface IXLPhonetic : IEquatable<IXLPhonetic>
{
    public string Text { get; }
    public int Start { get; }
    public int End { get; }
}
