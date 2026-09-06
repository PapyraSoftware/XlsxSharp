namespace XlsxSharp.IO.Packaging;

/// <summary>
/// The package does not conform to ECMA-376 Part 2 (Open Packaging Conventions), or an operation
/// on it would produce a package that does not.
/// </summary>
public class OpcException : Exception
{
    public OpcException()
        : base("The package is not a valid Open Packaging Conventions package.") { }

    public OpcException(string message)
        : base(message) { }

    public OpcException(string message, Exception innerException)
        : base(message, innerException) { }

    internal static OpcException InvalidPartName(string partName, string reason) =>
        new($"'{partName}' is not a valid part name, because {reason}.");

    internal static OpcException PartNotFound(string partName) =>
        new($"The package has no part named '{partName}'.");

    internal static OpcException DuplicatePart(string partName) =>
        new($"The package already has a part named '{partName}'.");

    internal static OpcException RelationshipNotFound(string sourceName, string id) =>
        new($"The relationship '{id}' does not exist in {DescribeSource(sourceName)}.");

    internal static OpcException DuplicateRelationshipId(string sourceName, string id) =>
        new($"The relationship id '{id}' is already used in {DescribeSource(sourceName)}.");

    internal static OpcException ExternalRelationship(string sourceName, string id) =>
        new(
            $"The relationship '{id}' in {DescribeSource(sourceName)} targets an external "
                + "resource, which is not a part of the package."
        );

    internal static OpcException NoContentType(string partName) =>
        new(
            $"The part '{partName}' has no content type. [Content_Types].xml declares neither a "
                + "default for its extension nor an override for it."
        );

    private static string DescribeSource(string sourceName) =>
        sourceName.Length == 0 ? "the package" : $"the part '{sourceName}'";
}
