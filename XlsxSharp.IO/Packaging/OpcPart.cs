using System.IO.Compression;

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// A part of an <see cref="OpcPackage"/>: a named stream with a content type and its own
/// relationships (ECMA-376 Part 2 §9.1).
/// </summary>
/// <remarks>
/// Content is kept where it already is for as long as possible. A part read from a package keeps
/// pointing at its ZIP entry and is only materialised when someone writes to it, so opening a
/// workbook does not copy the parts it never touches.
/// </remarks>
public sealed class OpcPart
{
    private readonly OpcPackage _package;

    /// <summary>The entry this part was read from, or <c>null</c> for a part added in memory.</summary>
    private ZipArchiveEntry? _entry;

    /// <summary>The content after it was written to, or <c>null</c> while it still is in the ZIP.</summary>
    private MemoryStream? _content;

    internal OpcPart(
        OpcPackage package,
        string name,
        string contentType,
        ZipArchiveEntry? entry,
        MemoryStream? content
    )
    {
        this._package = package;
        this._entry = entry;
        this._content = content;
        this.Name = name;
        this.ContentType = contentType;
        this.Relationships = new OpcRelationshipCollection(name);
    }

    /// <summary>The absolute part name, e.g. <c>/xl/workbook.xml</c>.</summary>
    public string Name { get; }

    /// <summary>The content type declared for this part in <c>[Content_Types].xml</c>.</summary>
    public string ContentType { get; }

    /// <summary>The relationships declared by this part.</summary>
    public OpcRelationshipCollection Relationships { get; internal set; }

    /// <summary>
    /// Opens the content for reading. The returned stream is positioned at the start and is the
    /// caller's to dispose.
    /// </summary>
    public Stream GetReadStream()
    {
        this._package.ThrowIfDisposed();

        if (this._content is not null)
        {
            return new MemoryStream(
                this._content.GetBuffer(),
                0,
                (int)this._content.Length,
                writable: false
            );
        }

        if (this._entry is not null)
        {
            return this._entry.Open();
        }

        return new MemoryStream([], writable: false);
    }

    /// <summary>
    /// Opens the content for writing, discarding whatever the part held before. The returned
    /// stream is the caller's to dispose, and the part keeps what was written to it.
    /// </summary>
    public Stream GetWriteStream()
    {
        this._package.ThrowIfDisposed();
        this._package.ThrowIfReadOnly();

        this._entry = null;
        this._content = new MemoryStream();
        return new OpcPartWriteStream(this._content);
    }

    /// <summary>
    /// The part this part's relationship <paramref name="id"/> points at.
    /// </summary>
    /// <exception cref="OpcException">
    /// There is no such relationship, it targets something outside the package, or the part it
    /// targets does not exist.
    /// </exception>
    public OpcPart GetRelatedPart(string id)
    {
        OpcRelationship relationship = this.Relationships.GetById(id);
        return this._package.ResolveRelatedPart(relationship);
    }

    /// <summary>
    /// The parts related from here with the given relationship type, in document order.
    /// </summary>
    public IEnumerable<OpcPart> GetRelatedParts(string relationshipType) =>
        this
            .Relationships.OfType(relationshipType)
            .Where(r => r.TargetMode == OpcTargetMode.Internal)
            .Select(this._package.ResolveRelatedPart);

    /// <summary>
    /// Copies the content of this part into <paramref name="archive"/> as a new entry.
    /// </summary>
    internal void WriteTo(ZipArchive archive, CompressionLevel compressionLevel)
    {
        ZipArchiveEntry entry = archive.CreateEntry(
            this.Name.TrimStart('/'),
            // The bitmaps of a workbook are already compressed; deflating them again costs time
            // and gains nothing.
            IsAlreadyCompressed(this.ContentType)
                ? CompressionLevel.NoCompression
                : compressionLevel
        );

        using Stream target = entry.Open();
        using Stream source = this.GetReadStream();
        source.CopyTo(target);
    }

    private static bool IsAlreadyCompressed(string contentType) =>
        contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
        && !contentType.EndsWith("bmp", StringComparison.OrdinalIgnoreCase)
        && !contentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Hands out a writable view of the part's buffer without letting the caller disposing the
    /// handle take the buffer away from the part.
    /// </summary>
    private sealed class OpcPartWriteStream(MemoryStream content) : Stream
    {
        public override bool CanRead => false;

        public override bool CanSeek => true;

        public override bool CanWrite => true;

        public override long Length => content.Length;

        public override long Position
        {
            get => content.Position;
            set => content.Position = value;
        }

        public override void Flush() => content.Flush();

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException("The part was opened for writing.");

        public override long Seek(long offset, SeekOrigin origin) => content.Seek(offset, origin);

        public override void SetLength(long value) => content.SetLength(value);

        public override void Write(byte[] buffer, int offset, int count) =>
            content.Write(buffer, offset, count);

        public override void Write(ReadOnlySpan<byte> buffer) => content.Write(buffer);

        public override void WriteByte(byte value) => content.WriteByte(value);

        protected override void Dispose(bool disposing)
        {
            // Deliberately does not dispose the underlying MemoryStream: it is the part's
            // content and outlives this handle.
            content.Flush();
            base.Dispose(disposing);
        }
    }
}
