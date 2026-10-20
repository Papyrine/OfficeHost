/// <summary>
/// An Extended GUID ([MS-FSSHTTPB] 2.2.1.7): a GUID and an unsigned 32-bit value, written in 1, 17, 18, 19 or 21 bytes.
/// The null Extended GUID (2.2.1.7.1) is the empty GUID with the value zero.
/// </summary>
readonly record struct ExtendedGuid(Guid Guid, uint Value)
{
    public static ExtendedGuid Null => default;

    public bool IsNull => Guid == Guid.Empty && Value == 0;

    /// <summary>
    /// The size of the shortest encoding, which is the one written (2.2.1.7.1 to 2.2.1.7.5).
    /// </summary>
    public int Size
    {
        get
        {
            if (IsNull)
            {
                return 1;
            }

            if (Value <= 0x1F)
            {
                return 17;
            }

            if (Value <= 0x3FF)
            {
                return 18;
            }

            if (Value <= 0x1FFFF)
            {
                return 19;
            }

            return 21;
        }
    }

    public override string ToString() => $"{{{Guid}}},{Value}";
}

/// <summary>
/// A Serial Number ([MS-FSSHTTPB] 2.2.1.9): a GUID and an unsigned 64-bit value, written in 1 byte when null (2.2.1.9.1)
/// and 25 otherwise (2.2.1.9.2).
/// </summary>
readonly record struct SerialNumber(Guid Guid, ulong Value)
{
    public static SerialNumber Null => default;

    public bool IsNull => Guid == Guid.Empty && Value == 0;

    public int Size
    {
        get
        {
            if (IsNull)
            {
                return 1;
            }

            return 25;
        }
    }

    public override string ToString() => $"{{{Guid}}},{Value}";
}

/// <summary>
/// A Cell ID ([MS-FSSHTTPB] 2.2.1.10): two Extended GUIDs.
/// </summary>
readonly record struct CellId(ExtendedGuid First, ExtendedGuid Second)
{
    public static CellId Null => default;

    public int Size => First.Size + Second.Size;
}

/// <summary>
/// A File Chunk Reference ([MS-FSSHTTPB] 2.2.1.2): the start and length of a part of a file.
/// </summary>
readonly record struct FileChunkReference(ulong Start, ulong Length)
{
    public int Size => CompactUInt64.Size(Start) + CompactUInt64.Size(Length);
}
