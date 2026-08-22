#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using XlsxSharp.Utils;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel.Protection;

internal class XLSheetProtection : IXLSheetProtection
{
    public XLSheetProtection(Algorithm algorithm)
    {
        this.Algorithm = algorithm;
        this.AllowedElements = XLSheetProtectionElements.SelectEverything;
    }

    public Algorithm Algorithm { get; internal set; }
    public XLSheetProtectionElements AllowedElements { get; set; }

    public Boolean IsPasswordProtected =>
        this.IsProtected && !String.IsNullOrEmpty(this.PasswordHash);
    public Boolean IsProtected { get; internal set; }

    internal String Base64EncodedSalt { get; set; }
    internal String PasswordHash { get; set; }
    internal UInt32 SpinCount { get; set; } = 100000;

    public IXLSheetProtection AllowElement(
        XLSheetProtectionElements element,
        Boolean allowed = true
    )
    {
        if (!allowed)
        {
            return this.DisallowElement(element);
        }

        this.AllowedElements |= element;
        return this;
    }

    public IXLSheetProtection AllowEverything() =>
        this.AllowElement(XLSheetProtectionElements.Everything);

    public IXLSheetProtection AllowNone()
    {
        this.AllowedElements = XLSheetProtectionElements.None;
        return this;
    }

    public object Clone() =>
        new XLSheetProtection(this.Algorithm)
        {
            IsProtected = this.IsProtected,
            PasswordHash = this.PasswordHash,
            SpinCount = this.SpinCount,
            Base64EncodedSalt = this.Base64EncodedSalt,
            AllowedElements = this.AllowedElements,
        };

    public XLSheetProtection CopyFrom(
        IXLElementProtection<XLSheetProtectionElements> sheetProtection
    )
    {
        if (sheetProtection is XLSheetProtection xlSheetProtection)
        {
            this.IsProtected = xlSheetProtection.IsProtected;
            this.Algorithm = xlSheetProtection.Algorithm;
            this.PasswordHash = xlSheetProtection.PasswordHash;
            this.SpinCount = xlSheetProtection.SpinCount;
            this.Base64EncodedSalt = xlSheetProtection.Base64EncodedSalt;
            this.AllowedElements = xlSheetProtection.AllowedElements;
        }
        return this;
    }

    public IXLSheetProtection DisallowElement(XLSheetProtectionElements element)
    {
        this.AllowedElements &= ~element;
        return this;
    }

    public IXLSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm) =>
        this.Protect(String.Empty, algorithm);

    public IXLSheetProtection Protect(XLSheetProtectionElements allowedElements) =>
        this.Protect(string.Empty, DefaultProtectionAlgorithm, allowedElements);

    public IXLSheetProtection Protect(
        Algorithm algorithm,
        XLSheetProtectionElements allowedElements
    ) => this.Protect(string.Empty, algorithm, allowedElements);

    public IXLSheetProtection Protect(
        String password,
        Algorithm algorithm = DefaultProtectionAlgorithm,
        XLSheetProtectionElements allowedElements = XLSheetProtectionElements.SelectEverything
    )
    {
        if (this.IsProtected)
        {
            throw new InvalidOperationException("The worksheet is already protected");
        }
        else
        {
            this.IsProtected = true;

            password = password ?? "";

            this.Algorithm = algorithm;
            this.Base64EncodedSalt = CryptographicAlgorithms.GenerateNewSalt(this.Algorithm);
            this.PasswordHash = CryptographicAlgorithms.GetPasswordHash(
                this.Algorithm,
                password,
                this.Base64EncodedSalt,
                this.SpinCount
            );
        }

        this.AllowedElements = allowedElements;

        return this;
    }

    public IXLSheetProtection Unprotect() => this.Unprotect(String.Empty);

    public IXLSheetProtection Unprotect(String password)
    {
        if (this.IsProtected)
        {
            if (this.PasswordHash.Length > 0 && string.IsNullOrEmpty(password))
            {
                throw new InvalidOperationException("The worksheet is password protected");
            }

            string hash = CryptographicAlgorithms.GetPasswordHash(
                this.Algorithm,
                password,
                this.Base64EncodedSalt,
                this.SpinCount
            );
            if (hash != this.PasswordHash)
            {
                throw new ArgumentException("Invalid password");
            }
            else
            {
                this.IsProtected = false;
                this.PasswordHash = String.Empty;
                this.Base64EncodedSalt = String.Empty;
            }
        }

        return this;
    }

    #region IXLProtectable interface

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowElement(
        XLSheetProtectionElements element,
        Boolean allowed
    ) => this.AllowElement(element, allowed);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowEverything() =>
        this.AllowEverything();

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowNone() =>
        this.AllowNone();

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.CopyFrom(
        IXLElementProtection<XLSheetProtectionElements> protectable
    ) => this.CopyFrom(protectable);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.DisallowElement(
        XLSheetProtectionElements element
    ) => this.DisallowElement(element);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Protect(
        Algorithm algorithm
    ) => this.Protect(algorithm);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Protect(
        String password,
        Algorithm algorithm
    ) => this.Protect(password, algorithm);

    IXLSheetProtection IXLSheetProtection.Protect(XLSheetProtectionElements allowedElements) =>
        this.Protect(allowedElements);

    IXLSheetProtection IXLSheetProtection.Protect(
        Algorithm algorithm,
        XLSheetProtectionElements allowedElements
    ) => this.Protect(algorithm, allowedElements);

    IXLSheetProtection IXLSheetProtection.Protect(
        String password,
        Algorithm algorithm,
        XLSheetProtectionElements allowedElements
    ) => this.Protect(password, algorithm, allowedElements);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Unprotect() =>
        this.Unprotect();

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Unprotect(
        String password
    ) => this.Unprotect(password);

    #endregion IXLProtectable interface
}
