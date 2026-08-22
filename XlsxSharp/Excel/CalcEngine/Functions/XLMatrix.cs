#nullable disable

using System;

namespace XlsxSharp.Excel.CalcEngine.Functions;

internal class XLMatrix
{
    public XLMatrix L;
    public XLMatrix U;
    public int cols;
    private readonly CalcContext _ctx;
    private double detOfP = 1;
    public double[,] mat;
    private int[] pi;
    public int rows;

    private XLMatrix(int iRows, int iCols, CalcContext ctx) // XLMatrix Class constructor
    {
        this.rows = iRows;
        this.cols = iCols;
        this._ctx = ctx;
        this.mat = new double[this.rows, this.cols];
    }

    public XLMatrix(double[,] arr, CalcContext ctx)
        : this(arr.GetLength(0), arr.GetLength(1), ctx)
    {
        int roCount = arr.GetLength(0);
        int coCount = arr.GetLength(1);
        for (int ro = 0; ro < roCount; ro++)
        {
            for (int co = 0; co < coCount; co++)
            {
                this.mat[ro, co] = arr[ro, co];
            }
        }
    }

    public double this[int iRow, int iCol] // Access this matrix as a 2D array
    {
        get => this.mat[iRow, iCol];
        set => this.mat[iRow, iCol] = value;
    }

    public bool IsSingular()
    {
        for (int row = 0; row < this.rows; row++)
        {
            for (int col = 0; col < this.cols; col++)
            {
                this._ctx.ThrowIfCancelled();
                double element = this.mat[row, col];
                if (double.IsNaN(element) || double.IsInfinity(element))
                {
                    return true;
                }
            }
        }

        return false;
    }

    public bool IsSquare() => (this.rows == this.cols);

    public void SetCol(XLMatrix v, int k)
    {
        for (int i = 0; i < this.rows; i++)
        {
            this.mat[i, k] = v[i, 0];
        }
    }

    public void MakeLU() // Function for LU decomposition
    {
        if (!this.IsSquare())
        {
            throw new InvalidOperationException("The matrix is not square!");
        }

        this.L = this.IdentityMatrix(this.rows, this.cols);
        this.U = this.Duplicate();

        this.pi = new int[this.rows];
        for (int i = 0; i < this.rows; i++)
        {
            this._ctx.ThrowIfCancelled();
            this.pi[i] = i;
        }

        int k0 = 0;

        for (int k = 0; k < this.cols - 1; k++)
        {
            double p = 0;
            for (int i = k; i < this.rows; i++) // find the row with the biggest pivot
            {
                this._ctx.ThrowIfCancelled();
                if (Math.Abs(this.U[i, k]) > p)
                {
                    p = Math.Abs(this.U[i, k]);
                    k0 = i;
                }
            }
            if (p == 0)
            {
                throw new InvalidOperationException("The matrix is singular!");
            }

            int pom1 = this.pi[k];
            this.pi[k] = this.pi[k0];
            this.pi[k0] = pom1; // switch two rows in permutation matrix

            double pom2;
            for (int i = 0; i < k; i++)
            {
                this._ctx.ThrowIfCancelled();
                pom2 = this.L[k, i];
                this.L[k, i] = this.L[k0, i];
                this.L[k0, i] = pom2;
            }

            if (k != k0)
            {
                this.detOfP *= -1;
            }

            for (int i = 0; i < this.cols; i++) // Switch rows in U
            {
                this._ctx.ThrowIfCancelled();
                pom2 = this.U[k, i];
                this.U[k, i] = this.U[k0, i];
                this.U[k0, i] = pom2;
            }

            for (int i = k + 1; i < this.rows; i++)
            {
                this.L[i, k] = this.U[i, k] / this.U[k, k];
                for (int j = k; j < this.cols; j++)
                {
                    this._ctx.ThrowIfCancelled();
                    this.U[i, j] = this.U[i, j] - this.L[i, k] * this.U[k, j];
                }
            }
        }
    }

    public XLMatrix SolveWith(XLMatrix v) // Function solves Ax = v in conformity with solution vector "v"
    {
        if (this.rows != this.cols)
        {
            throw new InvalidOperationException("The matrix is not square!");
        }

        if (this.rows != v.rows)
        {
            throw new ArgumentException("Wrong number of results in solution vector!");
        }

        if (this.L == null)
        {
            this.MakeLU();
        }

        XLMatrix b = new(this.rows, 1, this._ctx);
        for (int i = 0; i < this.rows; i++)
        {
            this._ctx.ThrowIfCancelled();
            b[i, 0] = v[this.pi[i], 0]; // switch two items in "v" due to permutation matrix
        }

        XLMatrix z = this.SubsForth(this.L, b);
        XLMatrix x = this.SubsBack(this.U, z);

        return x;
    }

    public XLMatrix Invert() // Function returns the inverted matrix
    {
        if (this.L == null)
        {
            this.MakeLU();
        }

        XLMatrix inv = new(this.rows, this.cols, this._ctx);

        for (int i = 0; i < this.rows; i++)
        {
            XLMatrix Ei = this.ZeroMatrix(this.rows, 1);
            Ei[i, 0] = 1;
            XLMatrix col = this.SolveWith(Ei);
            inv.SetCol(col, i);
        }
        return inv;
    }

    public double Determinant() // Function for determinant
    {
        if (this.L == null)
        {
            this.MakeLU();
        }

        double det = this.detOfP;
        for (int i = 0; i < this.rows; i++)
        {
            det *= this.U[i, i];
        }

        return det;
    }

    public XLMatrix Duplicate() // Function returns the copy of this matrix
    {
        XLMatrix matrix = new(this.rows, this.cols, this._ctx);
        for (int i = 0; i < this.rows; i++)
        {
            for (int j = 0; j < this.cols; j++)
            {
                this._ctx.ThrowIfCancelled();
                matrix[i, j] = this.mat[i, j];
            }
        }

        return matrix;
    }

    public XLMatrix SubsForth(XLMatrix A, XLMatrix b) // Function solves Ax = b for A as a lower triangular matrix
    {
        if (A.L == null)
        {
            A.MakeLU();
        }

        int n = A.rows;
        XLMatrix x = new(n, 1, this._ctx);

        for (int i = 0; i < n; i++)
        {
            x[i, 0] = b[i, 0];
            for (int j = 0; j < i; j++)
            {
                this._ctx.ThrowIfCancelled();
                x[i, 0] -= A[i, j] * x[j, 0];
            }
            x[i, 0] = x[i, 0] / A[i, i];
        }
        return x;
    }

    public XLMatrix SubsBack(XLMatrix A, XLMatrix b) // Function solves Ax = b for A as an upper triangular matrix
    {
        if (A.L == null)
        {
            A.MakeLU();
        }

        int n = A.rows;
        XLMatrix x = new(n, 1, this._ctx);

        for (int i = n - 1; i > -1; i--)
        {
            x[i, 0] = b[i, 0];
            for (int j = n - 1; j > i; j--)
            {
                this._ctx.ThrowIfCancelled();
                x[i, 0] -= A[i, j] * x[j, 0];
            }
            x[i, 0] = x[i, 0] / A[i, i];
        }
        return x;
    }

    public XLMatrix ZeroMatrix(int iRows, int iCols) // Function generates the zero matrix
    {
        XLMatrix matrix = new(iRows, iCols, this._ctx);
        for (int i = 0; i < iRows; i++)
        {
            for (int j = 0; j < iCols; j++)
            {
                this._ctx.ThrowIfCancelled();
                matrix[i, j] = 0;
            }
        }

        return matrix;
    }

    public XLMatrix IdentityMatrix(int iRows, int iCols) // Function generates the identity matrix
    {
        XLMatrix matrix = this.ZeroMatrix(iRows, iCols);
        for (int i = 0; i < Math.Min(iRows, iCols); i++)
        {
            this._ctx.ThrowIfCancelled();
            matrix[i, i] = 1;
        }

        return matrix;
    }
}
