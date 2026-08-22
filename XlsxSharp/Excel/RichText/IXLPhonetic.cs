using System;

namespace XlsxSharp.Excel.RichText;

public interface IXLPhonetic : IEquatable<IXLPhonetic>
{
    string Text { get; }
    int Start { get; }
    int End { get; }
}
