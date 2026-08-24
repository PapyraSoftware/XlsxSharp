#nullable disable

namespace XlsxSharp.Excel.Exceptions;

public class EmptyTableException : ClosedXMLException
{
    public EmptyTableException()
        : base() { }

    public EmptyTableException(string message)
        : base(message) { }

    public EmptyTableException(string message, Exception innerException)
        : base(message, innerException) { }
}
