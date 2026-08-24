using System.Diagnostics;
using System.Xml;

namespace XlsxSharp.IO;

/// <summary>
/// Markup Compatibility and Extensibility (MCE) processor per ISO-29500-3:2015.
/// </summary>
/// <remarks>
/// Does not process attributes. Fundamentally, it is a facade over <see cref="XmlReader"/> and
/// there is no benefit in skipping attributes. The consuming application asks for a presence of
/// an attribute. If consuming application won't process attribute, it won't even ask fo it.
/// </remarks>
public class MceXmlReader : IXmlReader
{
    private readonly XmlReader _reader;

    // MCE processes every element, so it must be fast. All element/attribute comparisons are done
    // with atomized strings from XmlReader name table.
    private readonly string _mce;
    private readonly XmlName _alternateContent;
    private readonly XmlName _choice;
    private readonly XmlName _fallback;
    private readonly XmlName _attRequires;
    private readonly XmlName _attIgnorable;
    private readonly XmlName _attProcessContent;
    private readonly XmlName _attMustUnderstand;

    private readonly Tracker<string> _ignorable = new(
        ReferenceEqualityComparer.Instance,
        static (state, ns) => state.ContainsKey(ns)
    );
    private readonly Tracker<NamePair> _processContent = new(
        EqualityComparer<NamePair>.Default,
        static (state, pair) =>
        {
            foreach (NamePair namePair in state.Keys)
            {
                if (namePair.Matches(pair))
                {
                    return true;
                }
            }

            return false;
        }
    );

    /// <summary>
    /// A set of namespaces understood by the consumer.
    /// </summary>
    private readonly HashSet<string> _appConfig = new(ReferenceEqualityComparer.Instance);

    /// <summary>
    /// An optional application defined extension element.
    /// </summary>
    private readonly XmlName? _adee;

    /// <summary>
    /// A handler to signal a mismatch.
    /// </summary>
    private readonly Action<MismatchInfo>? _signalMismatch;

    /// <summary>
    /// Is the reader currently in an application defined extension element? The value indicates
    /// depth at which we switched into ADEE mode.
    /// </summary>
    private int? _inAdee;

    public MceXmlReader(XmlReader reader, MceSettings settings)
    {
        const string mceNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        this._reader = reader;
        XmlNameTable nameTable =
            this._reader.NameTable ?? throw new ArgumentException("XmlReader must use name table.");
        this._mce = nameTable.Add(mceNs);
        this._alternateContent = XmlName.Atomize("AlternateContent", mceNs, nameTable);
        this._choice = XmlName.Atomize("Choice", mceNs, nameTable);
        this._fallback = XmlName.Atomize("Fallback", mceNs, nameTable);
        this._attRequires = XmlName.Atomize("Requires", "", nameTable);
        this._attIgnorable = XmlName.Atomize("Ignorable", mceNs, nameTable);
        this._attProcessContent = XmlName.Atomize("ProcessContent", mceNs, nameTable);
        this._attMustUnderstand = XmlName.Atomize("MustUnderstand", mceNs, nameTable);

        // Atomize MCE settings
        foreach (string appConfigNs in settings.ApplicationConfiguration)
        {
            this._appConfig.Add(nameTable.Add(appConfigNs));
        }

        if (settings.AdeeLocalName is { } extLocalName)
        {
            string atomizedExtName = nameTable.Add(extLocalName);
            string atomizedExtNs = nameTable.Add(settings.AdeeNamespaceUri ?? string.Empty);
            this._adee = new XmlName(atomizedExtName, atomizedExtNs);
        }

        this._signalMismatch = settings.SignalMismatch;
    }

    /// <inheritdoc/>
    public XmlTreeNodeType NodeType { get; private set; }

    /// <inheritdoc/>
    public int Depth => this._reader.Depth;

    /// <inheritdoc/>
    public string LocalName => this._reader.LocalName;

    /// <inheritdoc/>
    public string NamespaceUri => this._reader.NamespaceURI;

    /// <inheritdoc/>
    public string Value => this._reader.Value;

    /// <inheritdoc/>
    public LineInfo LineInfo => this._reader.GetLineInfo();

    /// <inheritdoc/>
    public bool Read()
    {
        if (this._inAdee is { } openedAdeeDepth)
        {
            if (!this.MoveToNextNode())
            {
                throw this.UnpairedXml();
            }

            if (this._reader.Depth == openedAdeeDepth)
            {
                this._inAdee = null;
            }

            return true;
        }

        // Loop should end when we are a normal element, not on MCE element. There can be nested AC inside
        while (this.MoveToNextNode())
        {
            if (this._adee is not null && this.IsOpenElement(this._adee.Value))
            {
                this._inAdee = this._reader.Depth;
                return true;
            }

            if (this.IsOpenElement(this._alternateContent))
            {
                this.SignalMismatchIfPresent();
                this.MoveToChoiceOrSkip();
            }
            else if (this.IsCloseElement(this._choice))
            {
                this.SkipToCloseAlternateContent(false);
            }
            else if (this.IsCloseElement(this._fallback))
            {
                this.SkipToCloseAlternateContent(true);
            }
            else if (this.IsIgnored())
            {
                // Ignored and not understood -> skip
                this.SkipToCloseElement();
            }
            else if (this.IsUnwrapped())
            {
                // The unwrapped open/close element was consumed and won't be emitted
                if (this.NodeType == XmlTreeNodeType.OpenElement)
                {
                    this.SignalMismatchIfPresent();
                }
            }
            else
            {
                // Not part of MCE, return the the consumer
                return true;
            }
        }

        return false;
    }

    /// <inheritdoc/>
    public string? GetAttribute(string attributeName, string? namespaceUri) =>
        // XmlReader returns to the element node once it reads the attribute value
        this._reader.GetAttribute(attributeName, namespaceUri ?? string.Empty);

    /// <inheritdoc/>
    public void Dispose() => this._reader.Dispose();

    private bool IsIgnored()
    {
        if (!this._ignorable.Declared(this._reader.NamespaceURI))
        {
            return false;
        }

        if (this._appConfig.Contains(this._reader.NamespaceURI))
        {
            return false;
        }

        if (
            this._processContent.Declared(
                new NamePair(this._reader.NamespaceURI, this._reader.LocalName)
            )
        )
        {
            return false;
        }

        return true;
    }

    private bool IsUnwrapped()
    {
        if (!this._ignorable.Declared(this._reader.NamespaceURI))
        {
            return false;
        }

        if (this._appConfig.Contains(this._reader.NamespaceURI))
        {
            return false;
        }

        if (
            !this._processContent.Declared(
                new NamePair(this._reader.NamespaceURI, this._reader.LocalName)
            )
        )
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// Move from <c>AlternateContent</c> open element to selected choice or to the closing
    /// element of <c>AlternateContent</c>. Does not emit nodes (even text ones), because
    /// specification says the <em>replace this AlternateContent element with the content of
    /// the Choice or Fallback element marked as selected.</em>.
    /// </summary>
    private void MoveToChoiceOrSkip()
    {
        Debug.Assert(this.IsOpenElement(this._alternateContent));

        while (this.MoveToNextNode())
        {
            if (this.IsCloseElement(this._alternateContent))
            {
                return;
            }

            if (this.IsOpenElement(this._choice))
            {
                if (
                    this.GetAttribute(this._attRequires) is not { } requires
                    || string.IsNullOrWhiteSpace(requires)
                )
                {
                    throw MceThrowHelper.InvalidAttribute(
                        this._attRequires.LocalName,
                        this._reader
                    );
                }

                if (this.IsChoiceSelected(requires))
                {
                    this.SignalMismatchIfPresent();
                    return;
                }

                this.SkipToCloseElement();
            }
            else if (this.IsOpenElement(this._fallback))
            {
                this.SignalMismatchIfPresent();
                return;
            }
            else if (this.NodeType == XmlTreeNodeType.OpenElement)
            {
                // AC should only contain only choice/fallback, but to future-proof, it can contain
                // other ignorable elements. Technically, it should also signal mismatch, but it's
                // illegal anyway so it makes no sense to signal mismatch.
                if (!this._ignorable.Declared(this._reader.NamespaceURI))
                {
                    throw MceThrowHelper.ElementNotIgnorable(this._reader.LocalName, this._reader);
                }

                this.SkipToCloseElement();
            }
        }

        throw this.UnpairedXml();
    }

    private void SkipToCloseAlternateContent(bool seenFallback)
    {
        Debug.Assert(this.IsCloseElement(this._choice) || this.IsCloseElement(this._fallback));
        int depth = this._reader.Depth;
        do
        {
            if (!this.MoveToNextNode())
            {
                throw this.UnpairedXml();
            }

            if (this._reader.Depth == depth && this.NodeType == XmlTreeNodeType.OpenElement)
            {
                if (this.IsElement(this._fallback))
                {
                    if (seenFallback)
                    {
                        throw MceThrowHelper.UnexpectedElementFound(
                            this._fallback.LocalName,
                            this._reader
                        );
                    }

                    seenFallback = true;
                }
                else if (this.IsElement(this._choice))
                {
                    if (seenFallback)
                    {
                        throw MceThrowHelper.UnexpectedElementFound(
                            this._choice.LocalName,
                            this._reader
                        );
                    }
                }
                else if (!this._ignorable.Declared(this._reader.NamespaceURI))
                {
                    throw MceThrowHelper.ElementNotIgnorable(this._reader.LocalName, this._reader);
                }
            }
        } while (this._reader.Depth >= depth);
    }

    private string? GetAttribute(XmlName attributeName)
    {
        if (!this._reader.HasAttributes)
        {
            return null;
        }

        // Once done, move back to the element container node, so various checks
        // on XmlReader.NodeType still work.
        while (this._reader.MoveToNextAttribute())
        {
            if (
                ReferenceEquals(this._reader.LocalName, attributeName.LocalName)
                && ReferenceEquals(this._reader.NamespaceURI, attributeName.Namespace)
            )
            {
                string value = this._reader.Value;
                this._reader.MoveToElement();
                return value;
            }
        }

        this._reader.MoveToElement();
        return null;
    }

    private bool IsChoiceSelected(string requires)
    {
        string[] prefixes = requires.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (string prefix in prefixes)
        {
            string? namespaceUri = this._reader.LookupNamespace(prefix);
            if (namespaceUri is null)
            {
                throw MceThrowHelper.NamespacePrefixNotFound(
                    this._attRequires.LocalName,
                    prefix,
                    this._reader
                );
            }

            if (!this._appConfig.Contains(namespaceUri))
            {
                return false;
            }
        }

        return true;
    }

    private bool IsOpenElement(XmlName name) =>
        this.NodeType == XmlTreeNodeType.OpenElement && this.IsElement(name);

    private bool IsCloseElement(XmlName name) =>
        this.NodeType == XmlTreeNodeType.CloseElement && this.IsElement(name);

    private bool IsElement(XmlName name)
    {
        Debug.Assert(this._reader.NameTable?.Get(name.LocalName) is not null);
        Debug.Assert(this._reader.NameTable?.Get(name.Namespace) is not null);
        return ReferenceEquals(this._reader.LocalName, name.LocalName)
            && ReferenceEquals(this._reader.NamespaceURI, name.Namespace);
    }

    private void SkipToCloseElement()
    {
        Debug.Assert(this.NodeType == XmlTreeNodeType.OpenElement);
        int depth = this._reader.Depth;
        do
        {
            if (!this.MoveToNextNode())
            {
                throw this.UnpairedXml();
            }
        } while (this._reader.Depth > depth);
    }

    /// <summary>
    /// Move to next opening or closing element from current element. This is the only permitted method that can move the reader from an element.
    /// </summary>
    private bool MoveToNextNode()
    {
        if (
            this._reader.NodeType == XmlNodeType.Element
            && this._reader.IsEmptyElement
            && this.NodeType == XmlTreeNodeType.OpenElement
        )
        {
            this.UntrackMceAttributes();
            this.NodeType = XmlTreeNodeType.CloseElement;
            return true;
        }

        while (this._reader.Read())
        {
            switch (this._reader.NodeType)
            {
                case XmlNodeType.Element:
                    this.TrackMceAttributes();
                    this.NodeType = XmlTreeNodeType.OpenElement;
                    return true;

                case XmlNodeType.EndElement:
                    this.UntrackMceAttributes();
                    this.NodeType = XmlTreeNodeType.CloseElement;
                    return true;

                // Text nodes:
                //   The Whitespace node should only appear if XmlReaderSetting.IgnoreWhitespace is
                //   set to false. The 'default' whitespace processing node should depend on the XML
                //   processor, if configuration says give me all whitespaces, give all whitespaces.
                case XmlNodeType.Whitespace:
                case XmlNodeType.Text:
                case XmlNodeType.SignificantWhitespace:
                case XmlNodeType.CDATA:
                case XmlNodeType.EntityReference:
                    this.NodeType = XmlTreeNodeType.Text;
                    return true;

                // Invalid nodes:
                //   We should never see these. If we do, there is a bug in the reader.
                //   None - If XmlReader.Read() returned false, the reader is never on None node
                //   Attribute - attribute reading code must ensure it ends back on the Element node
                case XmlNodeType.None:
                case XmlNodeType.Attribute:
                    throw new UnreachableException($"Encountered a node {this._reader.NodeType}.");

                // Skip nodes:
                //   Nodes that don't produce a XmlTreeNode. Depending on XmlReaderSetting, we might
                //   see them or not, but they don't produce node to consume in any case.
                case XmlNodeType.XmlDeclaration:
                case XmlNodeType.Comment:
                case XmlNodeType.ProcessingInstruction:
                case XmlNodeType.DocumentType:

                // Non-appearing:
                //    The documentation of the XmlReader.NodeType says the property never returns
                //    these types. They only appear in some in-memory XmlDocuments and enum was
                //    used in several places.
                case XmlNodeType.Document:
                case XmlNodeType.DocumentFragment:
                case XmlNodeType.Entity:
                case XmlNodeType.EndEntity:
                case XmlNodeType.Notation:
                default:
                    break;
            }
        }

        this.NodeType = XmlTreeNodeType.None;
        return false;
    }

    private void TrackMceAttributes()
    {
        Debug.Assert(this._reader.NodeType == XmlNodeType.Element);

        if (this.GetAttribute(this._attIgnorable) is { } ignorableValue)
        {
            this.TrackIgnorable(ignorableValue, this._attIgnorable);
        }

        if (this.GetAttribute(this._attProcessContent) is { } processContentValue)
        {
            this.TrackProcessableContent(processContentValue, this._attProcessContent);
        }
    }

    private void TrackIgnorable(string nsList, XmlName attribute)
    {
        foreach (string nsPrefix in nsList.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (this._reader.LookupNamespace(nsPrefix) is not { } ns)
            {
                throw MceThrowHelper.NamespacePrefixNotFound(
                    attribute.LocalName,
                    nsPrefix,
                    this._reader
                );
            }

            if (ReferenceEquals(ns, this._mce))
            {
                throw MceThrowHelper.MceNamespaceNotAllowed(attribute.LocalName, this._reader);
            }

            this._ignorable.Add(this._reader.Depth, ns);
        }
    }

    private void TrackProcessableContent(string processContent, XmlName attribute)
    {
        foreach (string token in processContent.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            NamePair namePair = NamePair.Parse(token, attribute, this._reader);
            if (ReferenceEquals(namePair.Namespace, this._mce))
            {
                throw MceThrowHelper.MceNamespaceNotAllowed(attribute.LocalName, this._reader);
            }

            if (!this._ignorable.Declared(namePair.Namespace))
            {
                throw MceThrowHelper.AttributeNamespaceNotIgnorable(
                    attribute.LocalName,
                    namePair.Namespace,
                    this._reader
                );
            }

            this._processContent.Add(this._reader.Depth, namePair);
        }
    }

    private void UntrackMceAttributes()
    {
        Debug.Assert(this._reader.NodeType is XmlNodeType.Element or XmlNodeType.EndElement);
        this._ignorable.Clear(this._reader.Depth);
        this._processContent.Clear(this._reader.Depth);
    }

    private void SignalMismatchIfPresent()
    {
        Debug.Assert(this.NodeType == XmlTreeNodeType.OpenElement);
        if (this._signalMismatch is null)
        {
            return;
        }

        if (this.GetAttribute(this._attMustUnderstand) is { } mustUnderstand)
        {
            foreach (
                string nsPrefix in mustUnderstand.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            )
            {
                if (this._reader.LookupNamespace(nsPrefix) is not { } ns)
                {
                    throw MceThrowHelper.NamespacePrefixNotFound(
                        this._attMustUnderstand.LocalName,
                        nsPrefix,
                        this._reader
                    );
                }

                if (!this._appConfig.Contains(ns))
                {
                    MismatchInfo info = new() { LineInfo = this._reader.GetLineInfo() };
                    this._signalMismatch(info);
                }
            }
        }
    }

    private Exception UnpairedXml() =>
        // An exception to throw when input is invalid XML (unpaired elements, ends without being
        // at the end of XML tree). That should never happen, because XmlReader should throw when
        // it detects an invalid XML.
        new UnreachableException(
            $"Not a valid XML stream (unpaired elements) at {this._reader.GetLineInfo()}."
        );

    /// <summary>
    /// A fully qualified XML name of element or an attribute.
    /// </summary>
    /// <param name="LocalName">Local name of an element.</param>
    /// <param name="Namespace">Default namespace is indicated by an empty string.</param>
    private readonly record struct XmlName(string LocalName, string Namespace)
    {
        internal static XmlName Atomize(
            string localName,
            string namespaceUri,
            XmlNameTable nameTable
        ) => new(nameTable.Add(localName), nameTable.Add(namespaceUri));
    };

    /// <summary>
    /// Tracker that keeps track of items encountered on the path from the root to the current
    /// element. It is used to check whether an item was declared on the current element or on
    /// an ancestor element. Because it must determine whether the item matches an item
    /// in the current element or <em>any</em> ancestor, it stores only the item found at
    /// the lowest depth, since items at higher depths are redundant.
    /// </summary>
    private class Tracker<T>
        where T : notnull
    {
        /// <summary>
        /// The key is a item, the value is first depth when it was encountered.
        /// </summary>
        private readonly Dictionary<T, int> _state;
        private readonly HashSet<int> _usedDepths;
        private readonly Func<Dictionary<T, int>, T, bool> _matches;

        internal Tracker(IEqualityComparer<T> comparer, Func<Dictionary<T, int>, T, bool> matches)
        {
            this._state = new Dictionary<T, int>(comparer);
            this._usedDepths = new HashSet<int>();
            this._matches = matches;
        }

        internal void Add(int depth, T item)
        {
            if (this._state.TryAdd(item, depth))
            {
                this._usedDepths.Add(depth);
            }
        }

        /// <summary>
        /// Was the matching value declared on the current element or an ancestor element?
        /// </summary>
        internal bool Declared(T value) => this._matches(this._state, value);

        /// <summary>
        /// Clear items from current item at depth.
        /// </summary>
        internal void Clear(int depthToClear)
        {
            if (!this._usedDepths.Remove(depthToClear))
            {
                return;
            }

            List<T> itemsToRemove = new();
            foreach ((T item, int depth) in this._state)
            {
                if (depthToClear == depth)
                {
                    itemsToRemove.Add(item);
                }
            }

            foreach (T item in itemsToRemove)
            {
                this._state.Remove(item);
            }
        }
    }

    /// <summary>
    /// Namespace - local name pair for processing content attribute.
    /// </summary>
    private readonly record struct NamePair(string Namespace, string? LocaleName)
    {
        internal static NamePair Parse(string token, XmlName attribute, XmlReader reader)
        {
            int commaIndex = token.IndexOf(':');
            if (commaIndex < 0)
            {
                throw MceThrowHelper.InvalidAttribute(reader.LocalName, reader);
            }

            string nsPrefix = token[..commaIndex];
            if (!IsValidName(nsPrefix))
            {
                throw MceThrowHelper.InvalidAttribute(attribute.LocalName, reader);
            }

            string nameToken = token[(commaIndex + 1)..];
            string? atomizedName;
            if (nameToken is ['*'])
            {
                atomizedName = null;
            }
            else
            {
                if (!IsValidName(nameToken))
                {
                    throw MceThrowHelper.InvalidAttribute(attribute.LocalName, reader);
                }

                atomizedName = reader.NameTable!.Add(nameToken);
            }

            if (reader.LookupNamespace(nsPrefix) is not { } ns)
            {
                throw MceThrowHelper.NamespacePrefixNotFound(attribute.LocalName, nsPrefix, reader);
            }

            return new NamePair(ns, atomizedName);
        }

        private static bool IsValidName(string nameToken) =>
            (nameToken.Length > 0 && nameToken.All(XmlConvert.IsNCNameChar));

        internal bool Matches(NamePair other)
        {
            if (ReferenceEquals(this.Namespace, other.Namespace))
            {
                if (ReferenceEquals(this.LocaleName, other.LocaleName))
                {
                    return true;
                }

                if (this.LocaleName is null || other.LocaleName is null)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
