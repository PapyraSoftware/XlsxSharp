namespace XlsxSharp.IO.Packaging;

/// <summary>
/// Navigating a SpreadsheetML package by part kind rather than by part name. This is what
/// replaces the SDK's <c>GetPartsOfType&lt;T&gt;</c> and <c>AddNewPart&lt;T&gt;</c>: the kind
/// carries the relationship type to look for and the content type to declare.
/// </summary>
public static class OoxmlPartExtensions
{
    /// <summary>The parts of the given kind related from the package, in relationship order.</summary>
    public static IEnumerable<OpcPart> PartsOfType(this OpcPackage package, OoxmlPartType partType)
    {
        ArgumentNullException.ThrowIfNull(package);
        ArgumentNullException.ThrowIfNull(partType);
        return package.GetRelatedParts(partType.RelationshipType);
    }

    /// <summary>The parts of the given kind related from a part, in relationship order.</summary>
    public static IEnumerable<OpcPart> PartsOfType(this OpcPart part, OoxmlPartType partType)
    {
        ArgumentNullException.ThrowIfNull(part);
        ArgumentNullException.ThrowIfNull(partType);
        return part.GetRelatedParts(partType.RelationshipType);
    }

    /// <summary>
    /// The single part of the given kind related from the package, or <c>null</c> when there is
    /// none.
    /// </summary>
    public static OpcPart? PartOfType(this OpcPackage package, OoxmlPartType partType) =>
        package.PartsOfType(partType).FirstOrDefault();

    /// <summary>
    /// The single part of the given kind related from a part, or <c>null</c> when there is none.
    /// </summary>
    public static OpcPart? PartOfType(this OpcPart part, OoxmlPartType partType) =>
        part.PartsOfType(partType).FirstOrDefault();

    /// <summary>
    /// Adds a part of the given kind to the package and relates the package to it.
    /// </summary>
    /// <param name="package">The package to add to.</param>
    /// <param name="partType">The kind of part to add.</param>
    /// <param name="contentType">
    /// Overrides the content type of the kind, for the parts whose type is per part rather than
    /// per kind, i.e. images.
    /// </param>
    /// <param name="partName">
    /// Overrides where the part goes. By default the next free name from the kind's template.
    /// </param>
    /// <param name="relationshipId">
    /// The relationship id to use, when the caller already has one - a save has to keep handing
    /// out the same ids it would have before, since a part's own content can carry another part's
    /// id as a plain attribute value. By default the next free id in the collection.
    /// </param>
    public static (OpcPart Part, string RelationshipId) AddPartOfType(
        this OpcPackage package,
        OoxmlPartType partType,
        string? contentType = null,
        string? partName = null,
        string? relationshipId = null
    )
    {
        ArgumentNullException.ThrowIfNull(package);
        ArgumentNullException.ThrowIfNull(partType);

        OpcPart part = AddPartCore(package, partType, contentType, partName);
        OpcRelationship relationship = package.Relationships.Add(
            part.Name,
            partType.RelationshipType,
            relationshipId
        );

        return (part, relationship.Id);
    }

    /// <summary>
    /// Adds a part of the given kind to the package and relates <paramref name="source"/> to it.
    /// </summary>
    /// <param name="source">The part that will point at the new one.</param>
    /// <param name="package">The package the new part is added to.</param>
    /// <param name="partType">The kind of part to add.</param>
    /// <param name="contentType">Overrides the content type of the kind.</param>
    /// <param name="partName">Overrides where the part goes.</param>
    /// <param name="relationshipId">The relationship id to use, by default the next free one.</param>
    public static (OpcPart Part, string RelationshipId) AddPartOfType(
        this OpcPart source,
        OpcPackage package,
        OoxmlPartType partType,
        string? contentType = null,
        string? partName = null,
        string? relationshipId = null
    )
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(package);
        ArgumentNullException.ThrowIfNull(partType);

        OpcPart part = AddPartCore(package, partType, contentType, partName);
        OpcRelationship relationship = source.Relationships.Add(
            part.Name,
            partType.RelationshipType,
            relationshipId
        );

        return (part, relationship.Id);
    }

    private static OpcPart AddPartCore(
        OpcPackage package,
        OoxmlPartType partType,
        string? contentType,
        string? partName
    ) =>
        package.AddPart(
            partName ?? NextFreePartName(package, partType),
            contentType ?? partType.ContentType
        );

    /// <summary>
    /// The first name from the kind's template that no part uses yet. Numbering starts at 1,
    /// which is the convention Excel follows.
    /// </summary>
    private static string NextFreePartName(OpcPackage package, OoxmlPartType partType)
    {
        if (!partType.IsNumbered)
        {
            return partType.PathTemplate;
        }

        for (int number = 1; ; number++)
        {
            string candidate = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                partType.PathTemplate,
                number
            );

            if (!package.TryGetPart(candidate, out _))
            {
                return candidate;
            }
        }
    }
}
