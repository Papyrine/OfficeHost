/// <summary>
/// The CRC-32 of zip files ([PKWARE-Zip] 4.4.7: the polynomial 0x04C11DB7, reflected), for the signatures of the data of
/// zip items ([MS-FSSHTTPD] 2.4.1).
/// </summary>
static class Crc32
{
    static readonly uint[] table = CreateTable();

    static uint[] CreateTable()
    {
        var result = new uint[256];
        for (var index = 0u; index < 256; index++)
        {
            var value = index;
            for (var bit = 0; bit < 8; bit++)
            {
                if ((value & 1) != 0)
                {
                    value = (value >> 1) ^ 0xEDB88320;
                }
                else
                {
                    value >>= 1;
                }
            }

            result[index] = value;
        }

        return result;
    }

    public static uint Compute(ReadOnlySpan<byte> data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var value in data)
        {
            crc = table[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        return ~crc;
    }
}
