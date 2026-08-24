using RBush;
using XlsxSharp.Excel;

namespace XlsxSharp.Utils;

/// <summary>
/// R-Tree for <see cref="Area"/> areas and some data associated with the area. The R-Tree
/// allows multiple occurrences per area.
/// </summary>
/// <typeparam name="TData">Data associated with each area.</typeparam>
internal sealed class RTree<TData>
{
    private readonly RBush<AreaData> _rBush = new();

    /// <summary>
    /// Number of nodes stores in R-Tree.
    /// </summary>
    internal int Count => this._rBush.Count;

    internal void Insert(Node node) => this._rBush.Insert(new AreaData(node.Area, node.Data));

    /// <remarks>
    /// It's not enough to specify only an area. The item also has to be specified, because area
    /// can contain multiple items (e.g. multiple conditional formats can be specified for same
    /// area).
    /// </remarks>
    internal void Delete(Node node) => this._rBush.Delete(new AreaData(node.Area, node.Data));

    internal List<Node> GetNodes(Area area, List<Node> buffer)
    {
        Envelope envelope = ToEnvelope(area);
        GetNodes(this._rBush.Root, in envelope, buffer);
        return buffer;
    }

    /// <summary>
    /// Info about leaf node of R-Tree.
    /// </summary>
    /// <param name="Area">Area of a sheet.</param>
    /// <param name="Data">Data associated with the <paramref name="Area"/> (e.g. color or hyperlink).</param>
    internal readonly record struct Node(Area Area, TData Data);

    // AreaData must implement Equals for RBush.Delete operation to properly work
    private sealed class AreaData : ISpatialData, IEquatable<AreaData>
    {
        private readonly Envelope _area;

        public AreaData(Area area, TData data)
        {
            this._area = ToEnvelope(area);
            this.Data = data;
        }

        public ref readonly Envelope Envelope => ref this._area;

        internal TData Data { get; }

        internal Area Area => ToArea(in this._area);

        public override int GetHashCode() =>
            // Class is never used in a hash table and equals is fast by itself.
            0;

        public override bool Equals(object? obj) => obj is AreaData other && this.Equals(other);

        public bool Equals(AreaData? other)
        {
            if (other is null)
            {
                return false;
            }

            if (ReferenceEquals(this, other))
            {
                return true;
            }

            // Use reference equals for data, because it might contain some business equals that
            // could mis-equal a data, e.g. border has business-like equals.
            return this._area.Equals(other._area) && ReferenceEquals(this.Data, other.Data);
        }
    }

    private static Envelope ToEnvelope(Area range) =>
        new(range.LeftColumn, range.TopRow, range.RightColumn, range.BottomRow);

    private static Area ToArea(in Envelope envelope) =>
        new((int)envelope.MinX, (int)envelope.MinY, (int)envelope.MaxX, (int)envelope.MaxY);

    private static void GetNodes(
        RBush<AreaData>.Node node,
        in Envelope boundingBox,
        List<Node> coveringNodes
    )
    {
        if (!node.Envelope.Contains(in boundingBox))
        {
            return;
        }

        if (node.IsLeaf)
        {
            for (int index = 0; index < node.Children.Count; index++)
            {
                AreaData castedChild = (AreaData)node.Children[index];
                if (castedChild.Envelope == boundingBox)
                {
                    coveringNodes.Add(new Node(castedChild.Area, castedChild.Data));
                }
            }

            return;
        }

        for (int index = 0; index < node.Children.Count; index++)
        {
            RBush<AreaData>.Node childNode = (RBush<AreaData>.Node)node.Children[index];
            GetNodes(childNode, in boundingBox, coveringNodes);
        }
    }
}
