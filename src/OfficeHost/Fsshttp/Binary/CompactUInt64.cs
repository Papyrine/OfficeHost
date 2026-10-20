/// <summary>
/// The compact unsigned 64-bit integer of [MS-FSSHTTPB] 2.2.1.1: 1 to 7 bytes for values up to 49 bits, with a type
/// flag in the low bits of the first byte, 9 bytes above.
/// </summary>
static class CompactUInt64
{
    /// <summary>
    /// The size of the shortest encoding of the value, which is the one written.
    /// </summary>
    public static int Size(ulong value)
    {
        if (value <= 0x7F)
        {
            return 1;
        }

        if (value <= 0x3FFF)
        {
            return 2;
        }

        if (value <= 0x1F_FFFF)
        {
            return 3;
        }

        if (value <= 0xFFF_FFFF)
        {
            return 4;
        }

        if (value <= 0x7_FFFF_FFFF)
        {
            return 5;
        }

        if (value <= 0x3FF_FFFF_FFFF)
        {
            return 6;
        }

        if (value <= 0x1_FFFF_FFFF_FFFF)
        {
            return 7;
        }

        return 9;
    }
}
