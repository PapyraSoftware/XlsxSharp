#nullable disable

using XlsxSharp.Excel.Drawings;
using XlsxSharp.Utils;

namespace XlsxSharp.Graphics;

internal class PngInfoReader : ImageInfoReader
{
    private const int CrcLength = 4;
    private const int SkippedHeaderLength = 5;

    private int[] MagicBytes { get; } = [137, 80, 78, 71, 13, 10, 26, 10];

    private const int HeaderType = 0x49484452; // IHDR
    private const int PhysicalDimensionType = 0x70485973; // pHYs

    protected override bool CheckHeader(Stream stream)
    {
        foreach (int magicByte in this.MagicBytes)
        {
            int streamByte = stream.ReadByte();
            if (streamByte != magicByte || streamByte == -1)
            {
                return false;
            }
        }
        return true;
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        stream.Position += this.MagicBytes.Length;
        uint hdrLength = stream.ReadU32BE();
        if (hdrLength != 13)
        {
            throw CorruptedException("Header length must be 13.");
        }

        if (ReadType(stream) != HeaderType)
        {
            throw CorruptedException("First chunk type must be IHDR.");
        }

        uint width = stream.ReadU32BE();
        uint height = stream.ReadU32BE();

        stream.Position += SkippedHeaderLength + CrcLength;

        uint pixelsPerUnitX = 0,
            pixelsPerUnitY = 0;
        while (stream.TryReadU32BE(out uint chunkLength))
        {
            uint chunkType = ReadType(stream);
            if (chunkType == PhysicalDimensionType)
            {
                pixelsPerUnitX = stream.ReadU32BE();
                pixelsPerUnitY = stream.ReadU32BE();
                byte unit = stream.ReadU8();
                bool isUnitMeter = unit == 1;
                if (!isUnitMeter)
                {
                    pixelsPerUnitX = pixelsPerUnitY = 0;
                }

                break;
            }

            stream.Position += chunkLength + CrcLength;
        }

        double dpiX = PixelsPerMeterToDpi(pixelsPerUnitX);
        double dpiY = PixelsPerMeterToDpi(pixelsPerUnitY);
        return new XLPictureInfo(XLPictureFormat.Png, width, height, dpiX, dpiY);
    }

    private static uint ReadType(Stream stream) => stream.ReadU32BE();

    private static ArgumentException CorruptedException(string text) =>
        new($"PNG is corrupted. {text}");

    private static double PixelsPerMeterToDpi(uint ppm) =>
        // Conversion from the common integer dots-per-inch to pixels-per-meter is lossy, so instead of 96 we get 95.9866
        ppm * 0.0254d;
}
