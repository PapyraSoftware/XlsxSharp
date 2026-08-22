namespace XlsxSharp.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A marker interface for elements that can be referenced from other parts of XSD.
/// </summary>
internal interface IParslet
{
    /// <summary>
    /// The name used by other to reference this element.
    /// </summary>
    public ParsletName Name { get; }

    /// <summary>
    /// Generate a <c>Parse*()</c> method that parses the referencable element.
    /// </summary>
    public void GenerateParseMethod(CodeBuilder code);
}
