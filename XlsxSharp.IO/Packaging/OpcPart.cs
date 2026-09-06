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
    private byte[]? _content;

    internal OpcPart(
        OpcPackage package,
        string name,
        string contentType,
        ZipArchiveEntry? entry,
        byte[]? content
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
    public string ContentType { get; internal set; }

    /// <summary>The relationships declared by this part.</summary>
    public OpcRelationshipCollection Relationships { get; internal set; }

    /// <summary>
    /// The uncompressed size of the part's content in bytes, without reading it. A part backed by
    /// a ZIP entry gets this straight from the entry's own directory record - unlike
    /// <see cref="GetReadStream"/>, which for such a part is a raw deflate stream and does not
    /// support <see cref="Stream.Length"/> at all.
    /// </summary>
    public long Length =>
        this._content is { } content ? content.Length
        : this._entry is { } entry ? entry.Length
        : 0;

    /// <summary>
    /// Opens the content for reading. The returned stream is positioned at the start and is the
    /// caller's to dispose.
    /// </summary>
    public Stream GetReadStream()
    {
        this._package.ThrowIfDisposed();

        if (this._content is not null)
        {
            return new MemoryStream(this._content, writable: false);
        }

        if (this._entry is not null)
        {
            return this._entry.Open();
        }

        return new MemoryStream([], writable: false);
    }

    /// <summary>
    /// Opens the content for writing, discarding whatever the part held before.
    /// </summary>
    /// <remarks>
    /// The returned stream is the caller's to dispose, and what was written reaches the part when
    /// that stream is flushed or disposed. Every writer over a part therefore has to be closed
    /// before the package is saved, which the <c>using</c> around it takes care of.
    /// </remarks>
    public Stream GetWriteStream()
    {
        this._package.ThrowIfDisposed();
        this._package.ThrowIfReadOnly();

        this._entry = null;
        this._content = [];
        return new OpcPartWriteStream(this);
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
    /// The part this part's relationship <paramref name="id"/> points at, or <c>null</c> when
    /// there is no such relationship, it targets something outside the package, or the part it
    /// targets does not exist.
    /// </summary>
    public OpcPart? GetRelatedPartOrDefault(string id) =>
        this.Relationships.TryGetById(id, out OpcRelationship? relationship)
        && relationship.TargetPartName is not null
        && this._package.TryGetPart(relationship.TargetPartName, out OpcPart? part)
            ? part
            : null;

    /// <summary>
    /// The parts related from here with the given relationship type, in document order. A
    /// dangling relationship - one whose target part does not actually exist in the package - is
    /// skipped rather than treated as an error; unlike <see cref="GetRelatedPart"/>, nothing here
    /// asks for one specific, expected-to-exist relationship by id.
    /// </summary>
    public IEnumerable<OpcPart> GetRelatedParts(string relationshipType) =>
        this
            .Relationships.OfType(relationshipType)
            .Where(r => r.TargetMode == OpcTargetMode.Internal)
            .Select(r =>
                r.TargetPartName is not null
                && this._package.TryGetPart(r.TargetPartName, out OpcPart? part)
                    ? part
                    : null
            )
            .OfType<OpcPart>();

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
    /// Buffers what is written to a part and hands it over on every flush, so that the part holds
    /// plain bytes rather than a stream it would have to own and dispose.
    /// </summary>
    private sealed class OpcPartWriteStream(OpcPart part) : Stream
    {
        private readonly MemoryStream content = new();

        public override bool CanRead => false;

        public override bool CanSeek => true;

        public override bool CanWrite => true;

        public override long Length => content.Length;

        public override long Position
        {
            get => content.Position;
            set => content.Position = value;
        }

        public override void Flush()
        {
            content.Flush();
            this.Publish();
        }

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
            if (disposing)
            {
                this.Publish();
                content.Dispose();
            }

            base.Dispose(disposing);
        }

        private void Publish() => part._content = content.ToArray();
    }
}
