namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// A non-state representation of a formula that can be used by many cells.
/// </summary>
internal class Formula
{
    public Formula(string text, ValueNode root)
    {
        this.AstRoot = root;
        this.Text = text;
    }

    /// <summary>Text of the formula.</summary>
    public string Text { get; }

    public ValueNode AstRoot { get; }
}
