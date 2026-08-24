using System.Diagnostics;
using System.Xml;
using static XlsxSharp.IO.XmlTreeNodeType;

namespace XlsxSharp.IO;

/// <summary>
/// <para>
/// Reader that expects that XML document consists of elements in a tree-like fashion. XML
/// shouldn't be mixed, text should be only in the leaves (e.g. <c>&lt;f&gt;ABS(A1)*A2&lt;/f&gt;</c>).
/// </para>
/// <para>
/// The schema of XML should be mostly by elements, with choices and sequences.
/// </para>
/// <para>
/// In case of some specialities or mixed content, use <c>XDocument.Load(_reader.ReadSubtree())</c>
/// and parse the result.
/// </para>
/// <para>
/// All <c>Get*</c> methods read values from attributes of current element.
/// </para>
/// <para>
/// The reader is always at either start element or end element. Any API that moves the reader will
/// end on either start or end element. The empty elements (e.g. &lt;br/&gt;) behave same way as
/// non-empty elements (that is different from <see cref="XmlReader"/> that checks
/// <see cref="XmlReader.IsEmptyElement"/>).
/// The reader element is used for one of two purposes:
/// <list type="bullet">
/// <item>
/// Element is being processed, i.e. parser logic has correctly identified the element (by
/// name) and will use parser logic to extract data from element (mostly by reading attributes,
/// potentially content if it is a leaf).
/// </item>
/// <item>
/// Element is used as a lookahead. The parsing logic is using it to determine how to parse the rest
/// of document. Example:
/// <example>Stylesheet parser has processed element <![CDATA[<name/>]]> from <![CDATA[<font><name val="Arial"/></font>]]>).
/// It reads next element <![CDATA[</font>]]> as a lookahead. The parsing logic now has to
/// determine what to do. Font could have additional properties (e.g. <![CDATA[<b/>]]>) or the
/// schema could end. Parser will use <c>reader.TryOpen("b")</c> to check if there is a bold
/// element. If there isn't, it uses <c>reader.TryClose("font")</c> to check that <c>font</c>
/// should close. If neither is true, XML doesn't match expected schema and parser will likely
/// throw an exception.
/// </example>
/// </item>
/// </list>
/// </para>
/// </summary>
public sealed class XmlTreeReader : IDisposable
{
    private static readonly XmlReaderSettings Settings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        Async = false,
        DtdProcessing = DtdProcessing.Prohibit,
        CloseInput = true,
    };

    /// <summary>
    /// The reader that holds the current element. The current node must always be either
    /// <see cref="OpenElement"/> or <see cref="CloseElement"/> after an API method call.
    /// </summary>
    private readonly IXmlReader _reader;

    private readonly IEnumMapper _enumMapper;

    /// <summary>
    /// What is current state of parser:
    /// <list type="bullet">
    /// <item>
    ///   <term>false</term>
    ///   <description>
    ///   Current element is being processed. we can get value of attributes.
    ///   </description>
    /// </item>
    /// <item>
    ///   <term>true</term>
    ///   <description>
    ///   Current element is not being processed. The only thing we are interested in is a name and
    ///   open/close. We are using it to determine how to parse the remainder of the file.
    ///   Trying to get attribute value will throw.
    ///   </description>
    /// </item>
    /// </list>
    /// </summary>
    private bool _inLookup;

    private readonly List<string> _context = new();

    public XmlTreeReader(Stream stream, IEnumMapper enumMapper, bool strictAttributeParsing)
    {
        this._reader = new MceXmlReader(
            XmlReader.Create(stream, Settings),
            new MceSettings
            {
                SignalMismatch = info =>
                    throw PartStructureException.MceError(
                        info.LineInfo,
                        "Mismatch between consuming application capability and document requirements."
                    ),
            }
        );
        this._enumMapper = enumMapper;
        this.StrictAttributeParsing = strictAttributeParsing;
    }

    /// <summary>
    /// Returns names of elements from currently processed element to root. The <c>On*Parsed</c>
    /// hook is called after processed element is closed. Therefore when the context is inspected
    /// in the <c>On*Parsed</c> hook, it doesn't contain the name of element that was just processed.
    /// </summary>
    public IReadOnlyList<string> Context => this._context;

    /// <summary>
    /// Should attributes with values outside of their simple type be treated as errors and throw
    /// an exception or be treated as missing and just return null? Excel generally ignores
    /// unparseable attributes.
    /// </summary>
    public bool StrictAttributeParsing { get; }

    /// <summary>
    /// Get the local name of current element (lookup/processing).
    /// </summary>
    internal string ElementName => this._reader.LocalName;

    internal LineInfo LineInfo => this._reader.LineInfo;

    /// <summary>
    /// Read next element. Check lookup element is <paramref name="localName"/>. If it is, open the
    /// element and return true. Otherwise, return false (element doesn't change).
    /// </summary>
    public bool TryOpen(string localName, string namespaceUri)
    {
        this.SwitchToLookup();
        Debug.Assert(this._reader.NodeType is OpenElement or CloseElement);

        if (
            this._reader.NodeType == OpenElement
            && this._reader.LocalName == localName
            && this._reader.NamespaceUri == namespaceUri
        )
        {
            // Element has been opened, so it should be processed.
            this.SwitchToProcessing();
            this._context.Add(localName);
            return true;
        }

        return false;
    }

    public bool TryClose(string localName, string namespaceUri)
    {
        this.SwitchToLookup();
        Debug.Assert(this._reader.NodeType is OpenElement or CloseElement);

        if (
            this._reader.NodeType == OpenElement
            || this._reader.LocalName != localName
            || this._reader.NamespaceUri != namespaceUri
        )
        {
            return false;
        }

        // Element has been closed, so it should be processed. Though closing elements are not
        // really processed, but we don't want to switch back to lookup. We just want to mark it
        // as "done." Be lazy, e.g. last element of a document doesn't have next element. If
        // parsing logic needs further elements, it will read them when they are needed.
        this.SwitchToProcessing();
        this._context.RemoveAt(this._context.Count - 1);
        return true;
    }

    /// <summary>
    /// Assert that we are at the element with <paramref name="localName"/>. Doesn't move anywhere.
    /// </summary>
    public void Open(string localName, string namespaceUri)
    {
        if (!this.TryOpen(localName, namespaceUri))
        {
            throw PartStructureException.ExpectedElementNotFound(
                $"Expected opening element '{localName}', but reader is currently on {this._reader.NodeType} node '{this._reader.LocalName}'.",
                this
            );
        }
    }

    /// <summary>
    /// Close the next unprocessed node. If the node doesn't match the <paramref name="localName"/>,
    /// throw an exception.
    /// </summary>
    public void Close(string localName, string namespaceUri)
    {
        if (!this.TryClose(localName, namespaceUri))
        {
            throw PartStructureException.ExpectedElementNotFound(
                $"Expected closing element '{localName}', but reader is currently on {this._reader.NodeType} node '{this._reader.LocalName}'.",
                this
            );
        }
    }

    /// <summary>
    /// Skip subtree that start on the current element. After subtree is read, the reader is
    /// on an ending element of a subtree in a processed state.
    /// </summary>
    /// <exception cref="InvalidOperationException">Reader isn't on opening element.</exception>
    public void Skip(string elementName)
    {
        this.ThrowOnNonStartElement();
        int startDepth = this._reader.Depth;
        do
        {
            if (!this._reader.Read())
            {
                throw this.InvalidXml();
            }
        } while (this._reader.NodeType != CloseElement || this._reader.Depth > startDepth);

        this._inLookup = false;
    }

    public bool? GetOptionalBool(string attributeName)
    {
        this.ThrowOnNonStartElement();
        bool? result = null;
        if (this._reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                result = XmlConvert.ToBoolean(value);
            }
            catch (FormatException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }
        }

        return result;
    }

    public int? GetOptionalInt(string attributeName)
    {
        this.ThrowOnNonStartElement();
        int? number = null;
        if (this._reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                number = XmlConvert.ToInt32(value);
            }
            catch (OverflowException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }
            catch (FormatException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }
        }

        return number;
    }

    public uint? GetOptionalUInt(string attributeName)
    {
        this.ThrowOnNonStartElement();
        long? number = null;
        if (this._reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                number = XmlConvert.ToInt64(value);
            }
            catch (OverflowException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }
            catch (FormatException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }

            if (number is < 0 or > uint.MaxValue)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value);
                }

                number = null;
            }
        }

        return number is not null ? (uint)number : null;
    }

    public double? GetOptionalDouble(string attributeName)
    {
        this.ThrowOnNonStartElement();
        double? number = null;
        if (this._reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                number = XmlConvert.ToDouble(value);
            }
            catch (OverflowException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }
            catch (FormatException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }

            if (
                number is not null
                && (double.IsNaN(number.Value) || double.IsInfinity(number.Value))
            )
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value);
                }

                number = null;
            }
        }

        return number;
    }

    /// <summary>
    /// Try to read <c>xsd:dateTime</c> from an attribute of current element.
    /// </summary>
    /// <param name="attributeName">Name of the attribute.</param>
    /// <returns>Read datetime or null if attribute is not present.</returns>
    public DateTime? GetOptionalDateTime(string attributeName)
    {
        this.ThrowOnNonStartElement();
        DateTime? dateTime = null;
        if (this._reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                dateTime = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind);
            }
            catch (FormatException e)
            {
                if (this.StrictAttributeParsing)
                {
                    this.ThrowAttributeFormatException(attributeName, value, e);
                }
            }
        }

        return dateTime;
    }

    public string? GetOptionalString(string attributeName)
    {
        this.ThrowOnNonStartElement();
        return this._reader.GetAttribute(attributeName, null);
    }

    public TEnum? GetOptionalEnum<TEnum>(string attributeName)
        where TEnum : struct, Enum
    {
        this.ThrowOnNonStartElement();
        string? enumString = this._reader.GetAttribute(attributeName, null);

        if (enumString is null)
        {
            return null;
        }

        if (!this._enumMapper.TryGetEnum<TEnum>(enumString, out TEnum enumValue))
        {
            if (this.StrictAttributeParsing)
            {
                this.ThrowAttributeFormatException(attributeName, enumString);
            }

            return null;
        }

        return enumValue;
    }

    public void Dispose() => this._reader.Dispose();

    private void SwitchToProcessing()
    {
        if (this._inLookup)
        {
            this._inLookup = false;
        }
    }

    private void SwitchToLookup()
    {
        // When switching to lookup, current node and all its attributes should have already been processed.
        if (this._inLookup)
        {
            return;
        }

        // Read next element.
        this.MoveToNextElement();
        this._inLookup = true;
    }

    /// <summary>
    /// Move from current opening/closing element to next opening/closing element.
    /// </summary>
    private void MoveToNextElement()
    {
        while (this._reader.Read())
        {
            if (this._reader.NodeType is OpenElement or CloseElement)
            {
                return;
            }
        }

        throw this.InvalidXml();
    }

    private void ThrowOnNonStartElement()
    {
        if (this._reader.NodeType != OpenElement || this._inLookup)
        {
            throw new InvalidOperationException(
                "To read content/attribute, the reader must be on start element and in processing state."
            );
        }
    }

    private void ThrowAttributeFormatException(
        string attributeName,
        string attributeValue,
        Exception? exception = null
    ) =>
        throw PartStructureException.InvalidAttributeFormat(
            attributeName,
            attributeValue,
            this,
            exception
        );

    private Exception InvalidXml() =>
        // This should never happen. The underlaying reader should throw when it detects
        // invalid XML (no root, unpaired elements, multiple elements in root)
        new UnreachableException($"{this._reader.LineInfo}: Invalid XML.");
}
