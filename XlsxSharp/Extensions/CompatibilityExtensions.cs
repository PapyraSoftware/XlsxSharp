#nullable disable

using System;
using System.IO;

// This file contains extensions methods that are present in .NET Core, but not in .NET Standard 2.0
#if !NETSTANDARD2_1_OR_GREATER
namespace XlsxSharp.Extensions
{
    internal static class StreamCompatibilityExtensions
    {
        public static int Read(this Stream stream, Span<byte> span)
        {
            for (int i = 0; i < span.Length; ++i)
            {
                int b = stream.ReadByte();
                if (b == -1)
                {
                    return i;
                }

                span[i] = (byte)b;
            }

            return span.Length;
        }
    }
}
#endif
