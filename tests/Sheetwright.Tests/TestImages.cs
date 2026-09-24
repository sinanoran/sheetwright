namespace Sheetwright.Tests;

/// <summary>
/// Image headers, hand-built, with dimensions the test chooses.
/// </summary>
/// <remarks>
/// <para>
/// Sheetwright reads an image's size out of its header rather than
/// decoding it, because a drawing anchor needs width and height in EMUs and
/// nothing else. That makes a header enough of an image for these tests, and it
/// keeps binary fixtures out of the repository — a checked-in PNG is a file
/// nobody can review.
/// </para>
/// <para>
/// It also lets a test say what the answer should be. A real photograph would
/// only prove the reader agrees with itself.
/// </para>
/// </remarks>
internal static class TestImages
{
    internal static byte[] Png(int width, int height)
    {
        byte[] image = new byte[33];
        byte[] signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        signature.CopyTo(image, 0);

        WriteBigEndian(image, 8, 13);
        "IHDR"u8.ToArray().CopyTo(image, 12);
        WriteBigEndian(image, 16, width);
        WriteBigEndian(image, 20, height);
        image[24] = 8;
        image[25] = 6;
        return image;
    }

    internal static byte[] Gif(int width, int height)
    {
        byte[] image = new byte[14];
        "GIF89a"u8.ToArray().CopyTo(image, 0);
        image[6] = (byte)(width & 0xFF);
        image[7] = (byte)(width >> 8);
        image[8] = (byte)(height & 0xFF);
        image[9] = (byte)(height >> 8);
        return image;
    }

    internal static byte[] Bmp(int width, int height)
    {
        byte[] image = new byte[54];
        image[0] = 0x42;
        image[1] = 0x4D;
        BitConverter.GetBytes(width).CopyTo(image, 18);
        BitConverter.GetBytes(height).CopyTo(image, 22);
        return image;
    }

    /// <summary>A JPEG whose only segment is the one carrying the dimensions.</summary>
    internal static byte[] Jpeg(int width, int height)
    {
        return
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            (byte)(height >> 8), (byte)(height & 0xFF),
            (byte)(width >> 8), (byte)(width & 0xFF),
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x01,
            0x03, 0x11, 0x01,
            0xFF, 0xD9
        ];
    }

    /// <summary>A little-endian TIFF with the two tags the reader looks for.</summary>
    internal static byte[] Tiff(int width, int height)
    {
        byte[] image = new byte[38];
        image[0] = 0x49;
        image[1] = 0x49;
        image[2] = 42;
        image[4] = 8;

        image[8] = 2;

        WriteEntry(image, 10, tag: 256, value: width);
        WriteEntry(image, 22, tag: 257, value: height);
        return image;

        static void WriteEntry(byte[] image, int offset, ushort tag, int value)
        {
            BitConverter.GetBytes(tag).CopyTo(image, offset);
            // Type 4: a four-byte unsigned integer, held inline because the count is one.
            BitConverter.GetBytes((ushort)4).CopyTo(image, offset + 2);
            BitConverter.GetBytes(1U).CopyTo(image, offset + 4);
            BitConverter.GetBytes((uint)value).CopyTo(image, offset + 8);
        }
    }

    private static void WriteBigEndian(byte[] target, int offset, int value)
    {
        target[offset] = (byte)(value >> 24);
        target[offset + 1] = (byte)(value >> 16);
        target[offset + 2] = (byte)(value >> 8);
        target[offset + 3] = (byte)value;
    }
}
