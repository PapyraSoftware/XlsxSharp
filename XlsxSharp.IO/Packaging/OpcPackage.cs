using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// An ECMA-376 Part 2 (Open Packaging Conventions) package: a ZIP archive of named parts, their
/// content types and the relationships between them.
/// </summary>
/// <remarks>
/// <para>
/// This is the piece XlsxSharp used to get from <c>DocumentFormat.OpenXml.Packaging</c>. It
/// deliberately implements only what OOXML needs: no digital signatures, no interleaved
/// (piece-wise) parts, no encryption.
/// </para>
/// <para>
/// A package opened for reading streams parts straight out of the ZIP. A package opened for
/// writing keeps modified parts in memory and writes the whole archive out on
/// <see cref="Save"/> or <see cref="Dispose"/>.
/// </para>
/// </remarks>
public sealed class OpcPackage : IDisposable
{
    private readonly Dictionary<string, OpcPart> _parts = new(OpcPartName.Comparer);
    private readonly OpcContentTypes _contentTypes;

    /// <summary>The archive parts are read from, or <c>null</c> for a package created from scratch.</summary>
    private readonly ZipArchive? _archive;

    /// <summary>The stream the archive was read from and that <see cref="Save"/> writes back to.</summary>
    private readonly Stream? _stream;

    /// <summary>Whether disposing the package also disposes <see cref="_stream"/>.</summary>
    private readonly bool _ownsStream;

    private OpcPackageProperties? _properties;

    private bool _disposed;

    private OpcPackage(
        ZipArchive? archive,
        Stream? stream,
        bool ownsStream,
        bool readOnly,
        OpcContentTypes contentTypes
    )
    {
        this._archive = archive;
        this._stream = stream;
        this._ownsStream = ownsStream;
        this._contentTypes = contentTypes;
        this.IsReadOnly = readOnly;
        this.Relationships = new OpcRelationshipCollection(string.Empty);
    }

    /// <summary>Whether the package refuses modification.</summary>
    public bool IsReadOnly { get; }

    /// <summary>The package level relationships, i.e. the content of <c>/_rels/.rels</c>.</summary>
    public OpcRelationshipCollection Relationships { get; private set; }

    /// <summary>Every part of the package, excluding the relationship parts.</summary>
    public IReadOnlyCollection<OpcPart> Parts => this._parts.Values;

    /// <summary>
    /// The core properties of the package. Reading them loads the part they live in; writing one
    /// makes <see cref="SaveTo"/> rewrite that part, creating it if the package has none.
    /// </summary>
    public OpcPackageProperties Properties => this._properties ??= this.ReadProperties();

    /// <summary>Opens a package from a file.</summary>
    /// <param name="path">The path of the package.</param>
    /// <param name="writable">
    /// When true the package can be modified and <see cref="Save"/> writes it back to
    /// <paramref name="path"/>. The file is read into memory first, so that saving does not have
    /// to read and write the same file at once.
    /// </param>
    public static OpcPackage Open(string path, bool writable = false)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);

        if (!writable)
        {
            FileStream file = new(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return OpenCore(file, ownsStream: true, readOnly: true, saveTo: null);
        }

        MemoryStream buffer = new();
        using (FileStream file = new(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            file.CopyTo(buffer);
        }

        buffer.Position = 0;
        return OpenCore(buffer, ownsStream: true, readOnly: false, saveTo: path);
    }

    /// <summary>
    /// Opens a package from a stream. The stream stays the caller's to dispose.
    /// </summary>
    /// <param name="stream">The stream to read the package from.</param>
    /// <param name="writable">
    /// When true the package can be modified and <see cref="Save"/> writes it back to
    /// <paramref name="stream"/>. The stream is read into memory first, exactly as
    /// <see cref="Open(string, bool)"/> does for a file, so that saving does not have to read and
    /// write the same stream at once - overwriting a live ZIP archive while parts still point into
    /// it would corrupt the package.
    /// </param>
    public static OpcPackage Open(Stream stream, bool writable = false)
    {
        ArgumentNullException.ThrowIfNull(stream);

        if (!writable)
        {
            return OpenCore(stream, ownsStream: false, readOnly: true, saveTo: null);
        }

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        MemoryStream buffer = new();
        stream.CopyTo(buffer);
        buffer.Position = 0;

        OpcPackage package = OpenCore(buffer, ownsStream: true, readOnly: false, saveTo: null);
        package.SaveToStream = stream;
        return package;
    }

    /// <summary>Creates an empty package that <see cref="Save"/> writes to <paramref name="path"/>.</summary>
    public static OpcPackage Create(string path)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        return new OpcPackage(
            archive: null,
            stream: null,
            ownsStream: false,
            readOnly: false,
            OpcContentTypes.CreateEmpty()
        )
        {
            SavePath = path,
        };
    }

    /// <summary>Creates an empty package in memory.</summary>
    public static OpcPackage Create() =>
        new(
            archive: null,
            stream: null,
            ownsStream: false,
            readOnly: false,
            OpcContentTypes.CreateEmpty()
        );

    /// <summary>Where <see cref="Dispose"/> writes the package to, when anywhere.</summary>
    private string? SavePath { get; set; }

    /// <summary>
    /// The stream <see cref="Dispose"/> writes the package back to, for a package opened writable
    /// from a stream. Never the same stream the package reads from - see
    /// <see cref="Open(Stream, bool)"/>.
    /// </summary>
    private Stream? SaveToStream { get; set; }

    /// <summary>Looks up a part by name.</summary>
    public bool TryGetPart(string partName, [NotNullWhen(true)] out OpcPart? part)
    {
        this.ThrowIfDisposed();
        return this._parts.TryGetValue(OpcPartName.Normalize(partName), out part);
    }

    /// <summary>Looks up a part by name.</summary>
    /// <exception cref="OpcException">The package has no such part.</exception>
    public OpcPart GetPart(string partName)
    {
        string name = OpcPartName.Normalize(partName);
        return this.TryGetPart(name, out OpcPart? part)
            ? part
            : throw OpcException.PartNotFound(name);
    }

    /// <summary>Adds an empty part.</summary>
    /// <exception cref="OpcException">A part of that name already exists.</exception>
    public OpcPart AddPart(string partName, string contentType)
    {
        this.ThrowIfDisposed();
        this.ThrowIfReadOnly();
        ArgumentException.ThrowIfNullOrEmpty(contentType);

        string name = OpcPartName.Normalize(partName);
        if (this._parts.ContainsKey(name))
        {
            throw OpcException.DuplicatePart(name);
        }

        if (OpcPartName.IsRelationshipPart(name))
        {
            throw OpcException.InvalidPartName(
                partName,
                "relationship parts are managed by the package itself"
            );
        }

        OpcPart part = new(this, name, contentType, entry: null, content: []);
        this._parts.Add(name, part);
        this._contentTypes.SetContentType(name, contentType);
        return part;
    }

    /// <summary>
    /// Removes a part and every relationship pointing at it, from the package and from the other
    /// parts. Removing a part that does not exist is a no-op.
    /// </summary>
    public void DeletePart(string partName)
    {
        this.ThrowIfDisposed();
        this.ThrowIfReadOnly();

        string name = OpcPartName.Normalize(partName);
        if (!this._parts.Remove(name))
        {
            return;
        }

        this._contentTypes.Remove(name);
        this.Relationships.RemoveTargetsOf(name);
        foreach (OpcPart part in this._parts.Values)
        {
            part.Relationships.RemoveTargetsOf(name);
        }
    }

    /// <summary>The part a package level relationship points at.</summary>
    public OpcPart GetRelatedPart(string id) =>
        this.ResolveRelatedPart(this.Relationships.GetById(id));

    /// <summary>The parts related from the package with the given relationship type.</summary>
    public IEnumerable<OpcPart> GetRelatedParts(string relationshipType) =>
        this
            .Relationships.OfType(relationshipType)
            .Where(r => r.TargetMode == OpcTargetMode.Internal)
            .Select(this.ResolveRelatedPart);

    /// <summary>Writes the package out. Only valid for a package that knows where to write to.</summary>
    public void Save()
    {
        this.ThrowIfDisposed();
        this.ThrowIfReadOnly();

        if (this.SaveToStream is { } stream)
        {
            stream.Position = 0;
            stream.SetLength(0);
            this.SaveTo(stream);
            return;
        }

        if (this.SavePath is null)
        {
            throw new InvalidOperationException(
                "This package has no path to save to. Use SaveTo(Stream) instead."
            );
        }

        // Write to a temporary file and move it into place, so that a failure half way through
        // does not leave a truncated workbook behind.
        string directory = Path.GetDirectoryName(Path.GetFullPath(this.SavePath))!;
        string temporary = Path.Combine(directory, $".{Guid.NewGuid():N}.tmp");

        try
        {
            using (
                FileStream file = new(
                    temporary,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None
                )
            )
            {
                this.SaveTo(file);
            }

            File.Move(temporary, this.SavePath, overwrite: true);
        }
        catch
        {
            File.Delete(temporary);
            throw;
        }
    }

    /// <summary>Writes the package to a stream. The stream stays open.</summary>
    public void SaveTo(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        this.ThrowIfDisposed();
        this.FlushProperties();

        using ZipArchive archive = new(stream, ZipArchiveMode.Create, leaveOpen: true);

        // [Content_Types].xml has to be the first entry: consumers, Excel included, look at it
        // before they look at anything else.
        ZipArchiveEntry contentTypes = archive.CreateEntry(
            OpcContentTypes.PartName.TrimStart('/'),
            CompressionLevel.Optimal
        );

        using (Stream contentTypesStream = contentTypes.Open())
        {
            this._contentTypes.Write(contentTypesStream);
        }

        WriteRelationships(archive, string.Empty, this.Relationships);

        foreach (OpcPart part in this._parts.Values.OrderBy(p => p.Name, OpcPartName.Comparer))
        {
            part.WriteTo(archive, CompressionLevel.Optimal);
            WriteRelationships(archive, part.Name, part.Relationships);
        }
    }

    public void Dispose()
    {
        if (this._disposed)
        {
            return;
        }

        try
        {
            if (!this.IsReadOnly && (this.SavePath is not null || this.SaveToStream is not null))
            {
                this.Save();
            }
        }
        finally
        {
            this._disposed = true;
            this._archive?.Dispose();
            if (this._ownsStream)
            {
                this._stream?.Dispose();
            }
        }
    }

    /// <summary>
    /// Loads the core properties from the part related from the package, or starts an empty set
    /// when the package has none.
    /// </summary>
    private OpcPackageProperties ReadProperties()
    {
        OpcRelationship? relationship = this
            .Relationships.OfType(OpcRelationshipType.CoreProperties)
            .FirstOrDefault(r => r.TargetMode == OpcTargetMode.Internal);

        if (
            relationship?.TargetPartName is null
            || !this._parts.TryGetValue(relationship.TargetPartName, out OpcPart? part)
        )
        {
            return new OpcPackageProperties();
        }

        using Stream stream = part.GetReadStream();
        return OpcPackageProperties.Read(stream);
    }

    /// <summary>
    /// Writes the core properties back into their part, adding the part and the relationship to
    /// it when the package does not have them yet.
    /// </summary>
    private void FlushProperties()
    {
        if (this._properties is not { IsDirty: true })
        {
            return;
        }

        OpcRelationship? relationship = this
            .Relationships.OfType(OpcRelationshipType.CoreProperties)
            .FirstOrDefault(r => r.TargetMode == OpcTargetMode.Internal);

        OpcPart part;
        if (
            relationship?.TargetPartName is not null
            && this._parts.TryGetValue(relationship.TargetPartName, out OpcPart? existing)
        )
        {
            part = existing;
        }
        else
        {
            part = this.AddPart(
                OpcPackageProperties.DefaultPartName,
                OpcContentType.CoreProperties
            );

            this.Relationships.Add(part.Name, OpcRelationshipType.CoreProperties);
        }

        using Stream stream = part.GetWriteStream();
        this._properties.Write(stream);
    }

    internal OpcPart ResolveRelatedPart(OpcRelationship relationship)
    {
        if (relationship.TargetPartName is null)
        {
            throw OpcException.ExternalRelationship(relationship.SourcePartName, relationship.Id);
        }

        return this.GetPart(relationship.TargetPartName);
    }

    internal void ThrowIfDisposed() =>
        ObjectDisposedException.ThrowIf(this._disposed, typeof(OpcPackage));

    internal void ThrowIfReadOnly()
    {
        if (this.IsReadOnly)
        {
            throw new InvalidOperationException("The package was opened read-only.");
        }
    }

    private static OpcPackage OpenCore(
        Stream stream,
        bool ownsStream,
        bool readOnly,
        string? saveTo
    )
    {
        ZipArchive archive;
        try
        {
            archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        }
        catch (InvalidDataException e)
        {
            if (ownsStream)
            {
                stream.Dispose();
            }

            throw new OpcException("The package is not a ZIP archive.", e);
        }

        ZipArchiveEntry? contentTypesEntry = archive.GetEntry(
            OpcContentTypes.PartName.TrimStart('/')
        );

        if (contentTypesEntry is null)
        {
            archive.Dispose();
            if (ownsStream)
            {
                stream.Dispose();
            }

            throw new OpcException(
                "The package has no [Content_Types].xml and therefore is not an Open Packaging "
                    + "Conventions package."
            );
        }

        OpcContentTypes contentTypes;
        using (Stream contentTypesStream = contentTypesEntry.Open())
        {
            contentTypes = OpcContentTypes.Read(contentTypesStream);
        }

        OpcPackage package = new(archive, stream, ownsStream, readOnly, contentTypes)
        {
            SavePath = saveTo,
        };

        // First pass: every entry that is a part. Relationship parts are held back, they are not
        // parts in their own right and are attached to their owner in the second pass.
        List<ZipArchiveEntry> relationshipEntries = [];
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            // Directory entries carry no content and are not parts.
            if (entry.FullName.EndsWith('/'))
            {
                continue;
            }

            string name = "/" + entry.FullName;
            if (OpcPartName.Comparer.Equals(name, OpcContentTypes.PartName))
            {
                continue;
            }

            if (OpcPartName.IsRelationshipPart(name))
            {
                relationshipEntries.Add(entry);
                continue;
            }

            string normalized = OpcPartName.Normalize(name);
            string? contentType = contentTypes.GetContentType(normalized);
            if (contentType is null)
            {
                throw OpcException.NoContentType(normalized);
            }

            package._parts[normalized] = new OpcPart(
                package,
                normalized,
                contentType,
                entry,
                content: null
            );
        }

        // Second pass, once every part exists, so that a relationship can be attached to its owner.
        foreach (ZipArchiveEntry entry in relationshipEntries)
        {
            string ownerName = GetRelationshipOwner("/" + entry.FullName);

            OpcRelationshipCollection relationships;
            using (Stream relationshipStream = entry.Open())
            {
                relationships = OpcRelationshipCollection.Read(ownerName, relationshipStream);
            }

            if (ownerName.Length == 0)
            {
                package.Relationships = relationships;
            }
            else if (package._parts.TryGetValue(ownerName, out OpcPart? owner))
            {
                owner.Relationships = relationships;
            }

            // A .rels part for a part that does not exist is dangling. Excel writes those when
            // it deletes a part without cleaning up, so ignoring it is friendlier than throwing.
        }

        return package;
    }

    /// <summary>
    /// The part a relationship part belongs to, or an empty string for <c>/_rels/.rels</c>.
    /// </summary>
    private static string GetRelationshipOwner(string relationshipPartName)
    {
        // "/xl/_rels/workbook.xml.rels" -> "/xl/workbook.xml", "/_rels/.rels" -> "".
        int relsFolder = relationshipPartName.LastIndexOf(
            "/_rels/",
            StringComparison.OrdinalIgnoreCase
        );

        string folder = relationshipPartName[..relsFolder];
        string fileName = relationshipPartName[(relsFolder + "/_rels/".Length)..];
        string ownerFileName = fileName[..^".rels".Length];

        return ownerFileName.Length == 0 ? string.Empty : $"{folder}/{ownerFileName}";
    }

    private static void WriteRelationships(
        ZipArchive archive,
        string sourcePartName,
        OpcRelationshipCollection relationships
    )
    {
        if (relationships.IsEmpty)
        {
            return;
        }

        ZipArchiveEntry entry = archive.CreateEntry(
            OpcPartName.GetRelationshipPartName(sourcePartName).TrimStart('/'),
            CompressionLevel.Optimal
        );

        using Stream stream = entry.Open();
        relationships.Write(stream);
    }
}
