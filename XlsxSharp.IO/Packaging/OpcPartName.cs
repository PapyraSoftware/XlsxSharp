namespace XlsxSharp.IO.Packaging;

/// <summary>
/// Part names as defined by ECMA-376 Part 2 (Open Packaging Conventions), §9.1.1.
/// A part name is an absolute path with '/' separators, e.g. <c>/xl/worksheets/sheet1.xml</c>.
/// Names are compared case-insensitively but stored with the casing they were created with,
/// because that casing is what ends up in the ZIP entry and in relationship targets.
/// </summary>
public static class OpcPartName
{
    /// <summary>
    /// Comparer for part names. OPC requires case-insensitive comparison, and ASCII casing is
    /// enough: §9.1.1.1 restricts part name segments to pchar, which is ASCII.
    /// </summary>
    public static StringComparer Comparer => StringComparer.OrdinalIgnoreCase;

    /// <summary>
    /// Validates a part name and returns it with a leading slash.
    /// </summary>
    /// <exception cref="OpcException">The name is not a valid part name.</exception>
    public static string Normalize(string partName)
    {
        ArgumentNullException.ThrowIfNull(partName);

        if (partName.Length == 0)
        {
            throw OpcException.InvalidPartName(partName, "it is empty");
        }

        string name = partName[0] == '/' ? partName : "/" + partName;

        if (name.Length == 1)
        {
            throw OpcException.InvalidPartName(partName, "it is only a slash");
        }

        if (name[^1] == '/')
        {
            throw OpcException.InvalidPartName(partName, "it ends with a slash");
        }

        foreach (Range segmentRange in SegmentRanges(name))
        {
            ReadOnlySpan<char> segment = name.AsSpan()[segmentRange];
            if (segment.IsEmpty)
            {
                throw OpcException.InvalidPartName(partName, "it has an empty segment");
            }

            if (segment is "." or "..")
            {
                throw OpcException.InvalidPartName(partName, "it has a '.' or '..' segment");
            }

            if (segment[^1] == '.')
            {
                throw OpcException.InvalidPartName(partName, "a segment ends with a dot");
            }
        }

        return name;
    }

    /// <summary>
    /// Resolves a relationship target against the part that declares it. Targets are relative to
    /// the folder of the source part, per ECMA-376 Part 2 §9.3.
    /// </summary>
    /// <param name="sourcePartName">
    /// The part holding the relationship, or an empty string for package level relationships.
    /// </param>
    /// <param name="target">The raw target as written in the .rels part.</param>
    public static string ResolveTarget(string sourcePartName, string target)
    {
        ArgumentNullException.ThrowIfNull(target);

        if (target.Length == 0)
        {
            throw OpcException.InvalidPartName(target, "a relationship target is empty");
        }

        if (target[0] == '/')
        {
            return Normalize(target);
        }

        // The base is the folder containing the source part. Package level relationships
        // (/_rels/.rels) resolve against the package root.
        string baseFolder = sourcePartName.Length == 0 ? "/" : GetFolder(sourcePartName);

        List<string> segments = [.. baseFolder.Split('/', StringSplitOptions.RemoveEmptyEntries)];

        foreach (string segment in target.Split('/'))
        {
            switch (segment)
            {
                case "" or ".":
                    break;

                case "..":
                    if (segments.Count == 0)
                    {
                        throw OpcException.InvalidPartName(
                            target,
                            $"it escapes the package root relative to '{sourcePartName}'"
                        );
                    }

                    segments.RemoveAt(segments.Count - 1);
                    break;

                default:
                    segments.Add(segment);
                    break;
            }
        }

        return Normalize("/" + string.Join('/', segments));
    }

    /// <summary>
    /// Expresses <paramref name="targetPartName"/> relative to the folder of
    /// <paramref name="sourcePartName"/>, which is how relationship targets are written.
    /// </summary>
    /// <param name="sourcePartName">
    /// The part that will hold the relationship, or an empty string for package level relationships.
    /// </param>
    /// <param name="targetPartName">The absolute name of the part being related to.</param>
    public static string MakeRelativeTarget(string sourcePartName, string targetPartName)
    {
        string source = sourcePartName.Length == 0 ? "/" : GetFolder(sourcePartName);
        string target = Normalize(targetPartName);

        string[] sourceSegments = source
            .Trim('/')
            .Split('/', StringSplitOptions.RemoveEmptyEntries);
        string[] targetSegments = target.TrimStart('/').Split('/');

        int common = 0;
        while (
            common < sourceSegments.Length
            && common < targetSegments.Length - 1
            && Comparer.Equals(sourceSegments[common], targetSegments[common])
        )
        {
            common++;
        }

        IEnumerable<string> up = Enumerable.Repeat("..", sourceSegments.Length - common);
        return string.Join('/', up.Concat(targetSegments.Skip(common)));
    }

    /// <summary>
    /// The name of the part holding the relationships of <paramref name="partName"/>. For the
    /// package itself (an empty string) that is <c>/_rels/.rels</c>.
    /// </summary>
    public static string GetRelationshipPartName(string partName)
    {
        if (partName.Length == 0)
        {
            return "/_rels/.rels";
        }

        string name = Normalize(partName);
        int lastSlash = name.LastIndexOf('/');
        return $"{name[..lastSlash]}/_rels/{name[(lastSlash + 1)..]}.rels";
    }

    /// <summary>
    /// True when the part is a relationship part, which is not addressable as a part of its own.
    /// </summary>
    public static bool IsRelationshipPart(string partName) =>
        partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
        && partName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The extension of a part name without the dot, lower-cased, or an empty string when the
    /// part name has no extension.
    /// </summary>
    public static string GetExtension(string partName)
    {
        int lastSlash = partName.LastIndexOf('/');
        int lastDot = partName.LastIndexOf('.');

        // The dot has to be in the last segment, but it may start it: the extension of
        // "/_rels/.rels" is "rels", which is what the Default entry for relationship parts keys on.
        return lastDot > lastSlash ? partName[(lastDot + 1)..].ToLowerInvariant() : string.Empty;
    }

    /// <summary>
    /// The folder of a part name, with a trailing slash, e.g. <c>/xl/</c> for
    /// <c>/xl/workbook.xml</c>.
    /// </summary>
    private static string GetFolder(string partName)
    {
        string name = Normalize(partName);
        return name[..(name.LastIndexOf('/') + 1)];
    }

    private static IEnumerable<Range> SegmentRanges(string name)
    {
        // name starts with '/', so the first segment starts at 1.
        int start = 1;
        for (int i = 1; i <= name.Length; i++)
        {
            if (i == name.Length || name[i] == '/')
            {
                yield return new Range(start, i);
                start = i + 1;
            }
        }
    }
}
