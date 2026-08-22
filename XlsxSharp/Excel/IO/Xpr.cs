namespace XlsxSharp.Excel.IO;

/// <summary>
/// XML Parse Result. Indicated whether element has been successfully parsed or if it was a not.
/// The non-generic version is used when Parse* method doesn't return value.
/// </summary>
internal readonly struct Xpr
{
    private Xpr(bool success) => this.IsSuccess = success;

    /// <summary>
    /// The implicit casting operator doesn't work for interfaces. Therefore the codegen will use
    /// this method to convert an element parsing result into a <see cref="Xpr{T}"/> instead.
    /// </summary>
    /// <typeparam name="T">Type of value retrieved from an element.</typeparam>
    /// <param name="value">Value to store.</param>
    public static Xpr<T> From<T>(T value) => new(value);

    /// <summary>
    /// Was element successfully parsed.
    /// </summary>
    public bool IsSuccess { get; }

    public bool IsFail => !this.IsSuccess;

    /// <summary>
    /// <c>Parse*</c> method wasn't able to match this element.
    /// </summary>
    public static Xpr Fail() => new(false);

    /// <summary>
    /// A factory method to create failed <see cref="Xpr{T}"/>.
    /// </summary>
    public static Xpr<T> Fail<T>() => new();

    /// <summary>
    /// A factory method to create successful <see cref="Xpr"/>, indicating element was parsed.
    /// </summary>
    public static Xpr Success() => new(true);
}
