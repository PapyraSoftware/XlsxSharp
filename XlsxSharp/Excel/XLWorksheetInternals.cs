using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Excel;

internal sealed class XLWorksheetInternals : IDisposable
{
    private bool _disposed;

    internal required XLCellsCollection CellsCollection
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XLColumnsCollection ColumnsCollection
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XLRowsCollection RowsCollection
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XLRanges MergedRanges
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        init;
    }

    public void Dispose()
    {
        if (this._disposed)
        {
            return;
        }

        this.CellsCollection.ValueSlice.DereferenceSlice();
        this.CellsCollection.Clear();
        this.ColumnsCollection.Clear();
        this.RowsCollection.Clear();
        this.MergedRanges.RemoveAll();
        this._disposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (this._disposed)
        {
            throw new ObjectDisposedException(nameof(XLWorksheetInternals));
        }
    }
}
