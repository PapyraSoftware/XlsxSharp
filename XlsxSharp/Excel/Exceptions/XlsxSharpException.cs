#nullable disable

namespace XlsxSharp.Excel.Exceptions;

public abstract class XlsxSharpException : Exception
{
    protected XlsxSharpException()
        : base() { }

    protected XlsxSharpException(string message)
        : base(message) { }

    protected XlsxSharpException(string message, Exception innerException)
        : base(message, innerException) { }
}
