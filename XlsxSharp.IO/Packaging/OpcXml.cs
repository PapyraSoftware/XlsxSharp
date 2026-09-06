using System.Text;
using System.Xml;

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// XML settings shared by the parts the packaging layer writes itself.
/// </summary>
internal static class OpcXml
{
    /// <summary>
    /// UTF-8 without a byte order mark. <see cref="Encoding.UTF8"/> emits one, and OOXML parts
    /// are written without it.
    /// </summary>
    internal static readonly Encoding NoBomUtf8 = new UTF8Encoding(
        encoderShouldEmitUTF8Identifier: false
    );

    internal static XmlReaderSettings ReaderSettings { get; } =
        new()
        {
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,

            // A package is untrusted input. Resolving a DTD would let it pull in external
            // entities, so refuse one outright.
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };
}
