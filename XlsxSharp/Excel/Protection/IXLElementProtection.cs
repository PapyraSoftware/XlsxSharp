#nullable disable

using System;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel.Protection;

public interface IXLElementProtection<T> : IXLElementProtection
    where T : struct
{
    /// <summary>Gets or sets the elements that are allowed to be edited by the user, i.e. those that are not protected.</summary>
    /// <value>The allowed elements.</value>
    public T AllowedElements { get; set; }

    /// <summary>
    /// Adds the specified element to the list of allowed elements.
    /// Beware that if you pass through "None", this will have no effect.
    /// </summary>
    /// <param name="element">The element to add</param>
    /// <param name="allowed">Set to <c>true</c> to allow the element or <c>false</c> to disallow the element</param>
    /// <returns>The current protection instance</returns>
    public IXLElementProtection<T> AllowElement(T element, bool allowed = true);

    /// <summary>Allows all elements to be edited.</summary>
    public IXLElementProtection<T> AllowEverything();

    /// <summary>Allows no elements to be edited. Protects all elements.</summary>
    public IXLElementProtection<T> AllowNone();

    /// <summary>Copies all the protection settings from a different instance.</summary>
    /// <param name="protectable">The protectable.</param>
    public IXLElementProtection<T> CopyFrom(IXLElementProtection<T> protectable);

    /// <summary>
    /// Removes the element to the list of allowed elements.
    /// Beware that if you pass through "None", this will have no effect.
    /// </summary>
    /// <param name="element">The element to remove</param>
    /// <returns>The current protection instance</returns>
    public IXLElementProtection<T> DisallowElement(T element);

    /// <summary>Protects this instance without a password.</summary>
    /// <param name="algorithm">The algorithm.</param>
    public IXLElementProtection<T> Protect(Algorithm algorithm = DefaultProtectionAlgorithm);

    /// <summary>Protects this instance using the specified password and password hash algorithm.</summary>
    /// <param name="password">The password.</param>
    /// <param name="algorithm">The algorithm.</param>
    public IXLElementProtection<T> Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm
    );

    /// <summary>Unprotects this instance without a password.</summary>
    public IXLElementProtection<T> Unprotect();

    /// <summary>Unprotects this instance using the specified password.</summary>
    /// <param name="password">The password.</param>
    public IXLElementProtection<T> Unprotect(string password);
}

public interface IXLElementProtection : ICloneable
{
    /// <summary>Gets the algorithm used to hash the password.</summary>
    /// <value>The algorithm.</value>
    public Algorithm Algorithm { get; }

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
}
