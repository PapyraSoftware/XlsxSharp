#nullable disable

// Keep this file CodeMaid organised and cleaned
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel.Protection;

public interface IXLProtectable<TProtection, in TElement> : IXLProtectable
    where TProtection : IXLElementProtection<TElement>
    where TElement : struct
{
    public TProtection Protection { get; set; }

    /// <summary>Protects this instance without a password.</summary>
    public TProtection Protect(TElement allowedElements);

    /// <summary>Protects this instance without a password.</summary>
    public new TProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm);

    /// <summary>Protects this instance with the specified password, password hash algorithm and set elements that the user is allowed to change.</summary>
    /// <param name="algorithm">The algorithm.</param>
    /// <param name="allowedElements">The allowed elements.</param>
    public TProtection Protect(Algorithm algorithm, TElement allowedElements);

    /// <summary>Protects this instance using the specified password and password hash algorithm.</summary>
    /// <param name="password">The password.</param>
    /// <param name="algorithm">The algorithm.</param>
    public new TProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm
    );

    /// <summary>Protects this instance with the specified password, password hash algorithm and set elements that the user is allowed to change.</summary>
    /// <param name="password">The password.</param>
    /// <param name="algorithm">The algorithm.</param>
    /// <param name="allowedElements">The allowed elements.</param>
    public TProtection Protect(string password, Algorithm algorithm, TElement allowedElements);

    /// <summary>Unprotects this instance without a password.</summary>
    public new TProtection Unprotect();

    /// <summary>Unprotects this instance using the specified password.</summary>
    /// <param name="password">The password.</param>
    public new TProtection Unprotect(string password);
}

public interface IXLProtectable
{
    /// <summary>Gets a value indicating whether this instance is protected with a password.</summary>
    /// <value>
    ///   <c>true</c> if this instance is password protected; otherwise, <c>false</c>.
    /// </value>
    public bool IsPasswordProtected { get; }

    /// <summary>Gets a value indicating whether this instance is protected, either with or without a password.</summary>
    /// <value>
    ///   <c>true</c> if this instance is protected; otherwise, <c>false</c>.
    /// </value>
    public bool IsProtected { get; }

    /// <summary>Protects this instance without a password.</summary>
    public IXLElementProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm);

    /// <summary>Protects this instance using the specified password and password hash algorithm.</summary>
    /// <param name="password">The password.</param>
    /// <param name="algorithm">The algorithm.</param>
    public IXLElementProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm
    );

    /// <summary>Unprotects this instance without a password.</summary>
    public IXLElementProtection Unprotect();

    /// <summary>Unprotects this instance using the specified password.</summary>
    /// <param name="password">The password.</param>
    public IXLElementProtection Unprotect(string password);
}
