using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace XlsxSharp.Utils;

internal static partial class XmlEncoder
{
    [GeneratedRegex("_(x[\\dA-Fa-f]{4})_")]
    private static partial Regex xHHHHRegex { get; }

    public static string EncodeString(string encodeStr)
    {
        encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

        StringBuilder sb = new(encodeStr.Length);
        int len = encodeStr.Length;
        for (int i = 0; i < len; ++i)
        {
            char currentChar = encodeStr[i];
            if (XmlConvert.IsXmlChar(currentChar))
            {
                sb.Append(currentChar);
            }
            else if (i + 1 < len && XmlConvert.IsXmlSurrogatePair(encodeStr[i + 1], currentChar))
            {
                sb.Append(currentChar);
                sb.Append(encodeStr[++i]);
            }
            else
            {
                sb.Append(XmlConvert.EncodeName(currentChar.ToString()));
            }
        }

        return sb.ToString();
    }
}
