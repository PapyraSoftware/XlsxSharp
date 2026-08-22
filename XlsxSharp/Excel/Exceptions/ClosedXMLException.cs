#nullable disable

using System;

namespace XlsxSharp.Excel.Exceptions;

public abstract class ClosedXMLException : Exception
{
    protected ClosedXMLException()
        : base() { }

    protected ClosedXMLException(string message)
        : base(message) { }

    protected ClosedXMLException(string message, Exception innerException)
        : base(message, innerException) { }
}
