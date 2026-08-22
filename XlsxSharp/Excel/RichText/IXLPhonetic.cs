using System;

namespace XlsxSharp.Excel.RichText;

public interface IXLPhonetic : IEquatable<IXLPhonetic>
{
    String Text { get; }
    Int32 Start { get; }
    Int32 End { get; }
}
