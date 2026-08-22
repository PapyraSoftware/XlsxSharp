#nullable disable

using System;
using System.Security.Cryptography;
using System.Text;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Utils;

internal static class CryptographicAlgorithms
{
    public static string GenerateNewSalt(Algorithm algorithm)
    {
        if (RequiresSalt(algorithm))
        {
            return GetSalt();
        }
        else
        {
            return string.Empty;
        }
    }

    public static string GetPasswordHash(
        Algorithm algorithm,
        string password,
        string salt = "",
        uint spinCount = 0
    )
    {
        ArgumentNullException.ThrowIfNull(password);

        ArgumentNullException.ThrowIfNull(salt);

        if (password.Length == 0)
        {
            return "";
        }

        switch (algorithm)
        {
            case Algorithm.SimpleHash:
                return GetDefaultPasswordHash(password);

            case Algorithm.SHA512:
                return GetSha512PasswordHash(password, salt, spinCount);

            default:
                return string.Empty;
        }
    }

    public static string GetSalt(int length = 32)
    {
#pragma warning disable SYSLIB0023
        using (RNGCryptoServiceProvider random = new())
#pragma warning restore SYSLIB0023
        {
            byte[] salt = new byte[length];
            random.GetNonZeroBytes(salt);
            return Convert.ToBase64String(salt);
        }
    }

    public static bool RequiresSalt(Algorithm algorithm)
    {
        switch (algorithm)
        {
            case Algorithm.SimpleHash:
                return false;

            case Algorithm.SHA512:
                return true;

            default:
                return false;
        }
    }

    private static string GetDefaultPasswordHash(string password)
    {
        ArgumentNullException.ThrowIfNull(password);

        // http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
        // http://sc.openoffice.org/excelfileformat.pdf - 4.18.4
        // http://web.archive.org/web/20080906232341/http://blogs.infosupport.com/wouterv/archive/2006/11/21/Hashing-password-for-use-in-SpreadsheetML.aspx
        byte[] passwordCharacters = Encoding.ASCII.GetBytes(password);
        int hash = 0;
        if (passwordCharacters.Length > 0)
        {
            int charIndex = passwordCharacters.Length;

            while (charIndex-- > 0)
            {
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters[charIndex];
            }
            // Main difference from spec, also hash with charcount
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= passwordCharacters.Length;
            hash ^= (0x8000 | ('N' << 8) | 'K');
        }

        return Convert.ToString(hash, 16).ToUpperInvariant();
    }

    private static string GetSha512PasswordHash(string password, string salt, uint spinCount)
    {
        ArgumentNullException.ThrowIfNull(password);

        ArgumentNullException.ThrowIfNull(salt);

        byte[] saltBytes = Convert.FromBase64String(salt);
        byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
        byte[] bytes = [.. saltBytes, .. passwordBytes];

        byte[] hashedBytes;

        hashedBytes = SHA512.HashData(bytes);

        bytes = new byte[hashedBytes.Length + sizeof(uint)];
        for (uint i = 0; i < spinCount; i++)
        {
            byte[] le = BitConverter.GetBytes(i);
            Array.Copy(hashedBytes, bytes, hashedBytes.Length);
            Array.Copy(le, 0, bytes, hashedBytes.Length, le.Length);
            hashedBytes = SHA512.HashData(bytes);
        }

        return Convert.ToBase64String(hashedBytes);
    }
}
