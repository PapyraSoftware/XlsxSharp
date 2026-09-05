using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// The relationships declared by the package or by one of its parts, i.e. the content of one
/// <c>.rels</c> part (ECMA-376 Part 2 §9.3).
/// </summary>
public sealed class OpcRelationshipCollection : IReadOnlyCollection<OpcRelationship>
{
    private const string Ns = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly Dictionary<string, OpcRelationship> _byId = new(StringComparer.Ordinal);

    /// <summary>
    /// Insertion order, so that rewriting a package that was read keeps the original order
    /// instead of shuffling relationships around.
    /// </summary>
    private readonly List<OpcRelationship> _ordered = [];

    private int _nextId = 1;

    internal OpcRelationshipCollection(string sourcePartName) =>
        this.SourcePartName = sourcePartName;

    /// <summary>
    /// The part declaring these relationships, or an empty string for the package itself.
    /// </summary>
    public string SourcePartName { get; }

    public int Count => this._ordered.Count;

    /// <summary>True when nothing has to be written for this collection.</summary>
    internal bool IsEmpty => this._ordered.Count == 0;

    public IEnumerator<OpcRelationship> GetEnumerator() => this._ordered.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    /// <summary>Looks up a relationship by id.</summary>
    public bool TryGetById(string id, [NotNullWhen(true)] out OpcRelationship? relationship) =>
        this._byId.TryGetValue(id, out relationship);

    /// <summary>Looks up a relationship by id.</summary>
    /// <exception cref="OpcException">There is no relationship with that id.</exception>
    public OpcRelationship GetById(string id) =>
        this._byId.TryGetValue(id, out OpcRelationship? relationship)
            ? relationship
            : throw OpcException.RelationshipNotFound(this.SourcePartName, id);

    /// <summary>All relationships of a given type, in document order.</summary>
    public IEnumerable<OpcRelationship> OfType(string relationshipType) =>
        this._ordered.Where(r =>
            string.Equals(r.RelationshipType, relationshipType, StringComparison.Ordinal)
        );

    /// <summary>
    /// The id under which <paramref name="targetPartName"/> is related, or <c>null</c> when it
    /// is not related from here.
    /// </summary>
    public string? GetIdOfTarget(string targetPartName)
    {
        string normalized = OpcPartName.Normalize(targetPartName);
        foreach (OpcRelationship relationship in this._ordered)
        {
            if (OpcPartName.Comparer.Equals(relationship.TargetPartName, normalized))
            {
                return relationship.Id;
            }
        }

        return null;
    }

    /// <summary>
    /// Adds a relationship to a part of this package, generating an id when none is given.
    /// </summary>
    public OpcRelationship Add(string targetPartName, string relationshipType, string? id = null)
    {
        string target = OpcPartName.MakeRelativeTarget(
            this.SourcePartName,
            OpcPartName.Normalize(targetPartName)
        );

        return this.AddCore(target, relationshipType, OpcTargetMode.Internal, id);
    }

    /// <summary>
    /// Adds a relationship to something outside the package, e.g. a hyperlink.
    /// </summary>
    public OpcRelationship AddExternal(string target, string relationshipType, string? id = null) =>
        this.AddCore(target, relationshipType, OpcTargetMode.External, id);

    /// <summary>Removes a relationship. Removing one that does not exist is a no-op.</summary>
    public void Remove(string id)
    {
        if (this._byId.Remove(id, out OpcRelationship? relationship))
        {
            this._ordered.Remove(relationship);
        }
    }

    /// <summary>Removes every relationship pointing at the given part.</summary>
    internal void RemoveTargetsOf(string targetPartName)
    {
        string normalized = OpcPartName.Normalize(targetPartName);
        foreach (OpcRelationship relationship in this._ordered.ToArray())
        {
            if (OpcPartName.Comparer.Equals(relationship.TargetPartName, normalized))
            {
                this.Remove(relationship.Id);
            }
        }
    }

    private OpcRelationship AddCore(
        string target,
        string relationshipType,
        OpcTargetMode targetMode,
        string? id
    )
    {
        ArgumentException.ThrowIfNullOrEmpty(relationshipType);

        if (id is null)
        {
            id = this.NextFreeId();
        }
        else if (this._byId.ContainsKey(id))
        {
            throw OpcException.DuplicateRelationshipId(this.SourcePartName, id);
        }

        OpcRelationship relationship = new(
            this.SourcePartName,
            id,
            relationshipType,
            target,
            targetMode
        );

        this._byId.Add(id, relationship);
        this._ordered.Add(relationship);
        return relationship;
    }

    private string NextFreeId()
    {
        string id;
        do
        {
            id = $"rId{this._nextId++}";
        } while (this._byId.ContainsKey(id));

        return id;
    }

    internal static OpcRelationshipCollection Read(string sourcePartName, Stream stream)
    {
        OpcRelationshipCollection relationships = new(sourcePartName);

        using XmlReader reader = XmlReader.Create(stream, OpcXml.ReaderSettings);
        while (reader.Read())
        {
            if (
                reader.NodeType != XmlNodeType.Element
                || reader.LocalName != "Relationship"
                || reader.NamespaceURI != Ns
            )
            {
                continue;
            }

            string? id = reader.GetAttribute("Id");
            string? type = reader.GetAttribute("Type");
            string? target = reader.GetAttribute("Target");
            if (id is null || type is null || target is null)
            {
                throw new OpcException(
                    $"A relationship in '{OpcPartName.GetRelationshipPartName(sourcePartName)}' "
                        + "is missing an Id, Type or Target attribute."
                );
            }

            OpcTargetMode targetMode = string.Equals(
                reader.GetAttribute("TargetMode"),
                "External",
                StringComparison.Ordinal
            )
                ? OpcTargetMode.External
                : OpcTargetMode.Internal;

            relationships.AddCore(target, type, targetMode, id);
        }

        return relationships;
    }

    internal void Write(Stream stream)
    {
        XmlWriterSettings settings = new() { CloseOutput = false, Encoding = OpcXml.NoBomUtf8 };

        using XmlWriter writer = XmlWriter.Create(stream, settings);
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("Relationships", Ns);

        foreach (OpcRelationship relationship in this._ordered)
        {
            writer.WriteStartElement("Relationship", Ns);
            writer.WriteAttributeString("Id", relationship.Id);
            writer.WriteAttributeString("Type", relationship.RelationshipType);
            writer.WriteAttributeString("Target", relationship.Target);
            if (relationship.TargetMode == OpcTargetMode.External)
            {
                writer.WriteAttributeString("TargetMode", "External");
            }

            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }
}
