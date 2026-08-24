namespace XlsxSharp.IO;

/// <summary>
/// Settings for <see cref="MceXmlReader"/>.
/// </summary>
public class MceSettings
{
    /// <summary>
    /// A set of namespaces understood by the consuming application.
    /// </summary>
    public IReadOnlyCollection<string> ApplicationConfiguration { get; init; } =
        Array.Empty<string>();

    /// <summary>
    /// A local name for application-defined extension element.
    /// </summary>
    public string? AdeeLocalName { get; init; }

    /// <summary>
    /// A namespace of application-defined extension element.
    /// </summary>
    public string? AdeeNamespaceUri { get; init; }

    /// <summary>
    /// A handler called when the MCE processor detects a mismatch.
    /// </summary>
    public Action<MismatchInfo>? SignalMismatch { get; init; }
}
