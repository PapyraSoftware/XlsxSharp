namespace XlsxSharp.IO.Packaging;

/// <summary>
/// Whether a relationship target is a part of the package or something outside of it
/// (ECMA-376 Part 2 §9.3).
/// </summary>
public enum OpcTargetMode
{
    /// <summary>The target is a part of this package.</summary>
    Internal,

    /// <summary>The target is a URI outside of the package, e.g. a hyperlink.</summary>
    External,
}

/// <summary>
/// A relationship from the package or from one of its parts to a target
/// (ECMA-376 Part 2 §9.3).
/// </summary>
public sealed class OpcRelationship
{
    internal OpcRelationship(
        string sourcePartName,
        string id,
        string relationshipType,
        string target,
        OpcTargetMode targetMode
    )
    {
        this.SourcePartName = sourcePartName;
        this.Id = id;
        this.RelationshipType = relationshipType;
        this.Target = target;
        this.TargetMode = targetMode;
        this.TargetPartName =
            targetMode == OpcTargetMode.Internal
                ? OpcPartName.ResolveTarget(sourcePartName, target)
                : null;
    }

    /// <summary>
    /// The part declaring the relationship, or an empty string for a package level relationship.
    /// </summary>
    public string SourcePartName { get; }

    /// <summary>The relationship id, unique within the source part.</summary>
    public string Id { get; }

    /// <summary>The relationship type URI.</summary>
    public string RelationshipType { get; }

    /// <summary>The target as written in the .rels part, relative to the source part's folder.</summary>
    public string Target { get; }

    /// <summary>Whether <see cref="Target"/> points inside or outside the package.</summary>
    public OpcTargetMode TargetMode { get; }

    /// <summary>
    /// The absolute name of the targeted part, or <c>null</c> when <see cref="TargetMode"/> is
    /// <see cref="OpcTargetMode.External"/>.
    /// </summary>
    public string? TargetPartName { get; }
}
