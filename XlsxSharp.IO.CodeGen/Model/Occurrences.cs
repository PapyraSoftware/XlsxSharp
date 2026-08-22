using System;

namespace XlsxSharp.IO.CodeGen.Model;

public readonly record struct Occurrences(int? Min, int? Max)
{
    public int ActualMin => this.Min ?? 1;

    public int ActualMax => this.Max ?? 1;

    internal bool HasFixedCount => this.ActualMin == this.ActualMax;

    internal ElementsCount Elements
    {
        get
        {
            int min = this.Min ?? 1;
            int max = this.Max ?? 1;
            return (min, max) switch
            {
                (0, 1) => ElementsCount.ZeroToOne,
                (0, int.MaxValue) => ElementsCount.ZeroToMany,
                (1, 1) => ElementsCount.OneToOne,
                (1, int.MaxValue) => ElementsCount.OneToMany,
                _ => throw new NotSupportedException(),
            };
        }
    }
};
