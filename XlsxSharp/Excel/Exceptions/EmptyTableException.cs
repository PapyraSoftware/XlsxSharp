#nullable disable

using System;

namespace XlsxSharp.Excel.Exceptions;

public class EmptyTableException : ClosedXMLException
{
    public EmptyTableException()
        : base() { }

    public EmptyTableException(String message)
        : base(message) { }

    public EmptyTableException(String message, Exception innerException)
        : base(message, innerException) { }
}
