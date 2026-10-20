using System.Text;

/// <summary>
/// Writes FSSHTTPB ([MS-FSSHTTPB] 2.2.1) to a buffer writer, in the shortest encodings: compact integers and Extended
/// GUIDs at their smallest, 16-bit stream object headers where the type and length allow it, 32-bit otherwise.
/// </summary>
ref struct FsshttpbWriter(IBufferWriter<byte> output)
{
    public void WriteByte(byte value)
    {
        output.GetSpan(1)[0] = value;
        output.Advance(1);
    }

    public void WriteUInt16(ushort value)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(output.GetSpan(2), value);
        output.Advance(2);
    }

    public void WriteUInt32(uint value)
    {
        BinaryPrimitives.WriteUInt32LittleEndian(output.GetSpan(4), value);
        output.Advance(4);
    }

    public void WriteUInt64(ulong value)
    {
        BinaryPrimitives.WriteUInt64LittleEndian(output.GetSpan(8), value);
        output.Advance(8);
    }

    public void WriteGuid(Guid value)
    {
        value.TryWriteBytes(output.GetSpan(16));
        output.Advance(16);
    }

    public void WriteBytes(ReadOnlySpan<byte> value) => output.Write(value);

    /// <summary>
    /// A compact unsigned 64-bit integer (2.2.1.1): the value shifted left by the size, under a type flag of one bit at
    /// the position of the size, or the type byte 128 and 8 bytes for values above 49 bits.
    /// </summary>
    public void WriteCompactUInt64(ulong value)
    {
        if (value == 0)
        {
            WriteByte(0);
            return;
        }

        var size = CompactUInt64.Size(value);
        if (size == 9)
        {
            WriteByte(0x80);
            WriteUInt64(value);
            return;
        }

        var encoded = (value << size) | (1UL << (size - 1));
        var span = output.GetSpan(size);
        for (var index = 0; index < size; index++)
        {
            span[index] = (byte) (encoded >> (8 * index));
        }

        output.Advance(size);
    }

    /// <summary>
    /// An Extended GUID (2.2.1.7), in the shortest of the 5, 10, 17 and 32 bit forms.
    /// </summary>
    public void WriteExtendedGuid(ExtendedGuid value)
    {
        if (value.IsNull)
        {
            WriteByte(0);
            return;
        }

        if (value.Guid == Guid.Empty)
        {
            throw new ArgumentException("An Extended GUID with a value has the empty GUID", nameof(value));
        }

        if (value.Value <= 0x1F)
        {
            WriteByte((byte) (0x04 | (value.Value << 3)));
        }
        else if (value.Value <= 0x3FF)
        {
            WriteUInt16((ushort) (0x20 | (value.Value << 6)));
        }
        else if (value.Value <= 0x1FFFF)
        {
            var encoded = 0x40 | (value.Value << 7);
            var span = output.GetSpan(3);
            span[0] = (byte) encoded;
            span[1] = (byte) (encoded >> 8);
            span[2] = (byte) (encoded >> 16);
            output.Advance(3);
        }
        else
        {
            WriteByte(0x80);
            WriteUInt32(value.Value);
        }

        WriteGuid(value.Guid);
    }

    /// <summary>
    /// A Serial Number (2.2.1.9).
    /// </summary>
    public void WriteSerialNumber(SerialNumber value)
    {
        if (value.IsNull)
        {
            WriteByte(0);
            return;
        }

        if (value.Guid == Guid.Empty)
        {
            throw new ArgumentException("A Serial Number with a value has the empty GUID", nameof(value));
        }

        WriteByte(0x80);
        WriteGuid(value.Guid);
        WriteUInt64(value.Value);
    }

    /// <summary>
    /// A Cell ID (2.2.1.10).
    /// </summary>
    public void WriteCellId(CellId value)
    {
        WriteExtendedGuid(value.First);
        WriteExtendedGuid(value.Second);
    }

    /// <summary>
    /// A File Chunk Reference (2.2.1.2).
    /// </summary>
    public void WriteFileChunkReference(FileChunkReference value)
    {
        WriteCompactUInt64(value.Start);
        WriteCompactUInt64(value.Length);
    }

    /// <summary>
    /// An Extended GUID Array (2.2.1.8).
    /// </summary>
    public void WriteExtendedGuidArray(IReadOnlyList<ExtendedGuid> values)
    {
        WriteCompactUInt64((ulong) values.Count);
        foreach (var value in values)
        {
            WriteExtendedGuid(value);
        }
    }

    /// <summary>
    /// A Cell ID Array (2.2.1.11).
    /// </summary>
    public void WriteCellIdArray(IReadOnlyList<CellId> values)
    {
        WriteCompactUInt64((ulong) values.Count);
        foreach (var value in values)
        {
            WriteCellId(value);
        }
    }

    /// <summary>
    /// A Binary Item (2.2.1.3).
    /// </summary>
    public void WriteBinaryItem(ReadOnlySpan<byte> value)
    {
        WriteCompactUInt64((ulong) value.Length);
        WriteBytes(value);
    }

    /// <summary>
    /// A String Item (2.2.1.4).
    /// </summary>
    public void WriteStringItem(string value)
    {
        WriteCompactUInt64((ulong) value.Length);
        WriteBytes(Encoding.Unicode.GetBytes(value));
    }

    /// <summary>
    /// A String Item Array (2.2.1.14).
    /// </summary>
    public void WriteStringItemArray(IReadOnlyList<string> values)
    {
        WriteCompactUInt64((ulong) values.Count);
        foreach (var value in values)
        {
            WriteStringItem(value);
        }
    }

    /// <summary>
    /// The start of a stream object (2.2.1.5.1, 2.2.1.5.2): 16-bit for a type below 0x40 and a length up to 127, 32-bit
    /// otherwise, with a Large Length from 32767. The compound bit comes from the type.
    /// </summary>
    public void WriteStart(StreamObjectType type, int length) => WriteStart(type, (ulong) length);

    public void WriteStart(StreamObjectType type, ulong length)
    {
        var compound = 0u;
        if (type.IsCompound())
        {
            compound = 0x4;
        }

        var code = (uint) type;
        if (code < 0x40 && length <= 127)
        {
            WriteUInt16((ushort) (compound | (code << 3) | ((uint) length << 9)));
            return;
        }

        if (length < 32767)
        {
            WriteUInt32(0x2 | compound | (code << 3) | ((uint) length << 17));
            return;
        }

        WriteUInt32(0x2 | compound | (code << 3) | (32767u << 17));
        WriteCompactUInt64(length);
    }

    /// <summary>
    /// The end of a compound stream object: 8-bit for a type below 0x40 (2.2.1.5.3), 16-bit otherwise (2.2.1.5.4).
    /// </summary>
    public void WriteEnd(StreamObjectType type)
    {
        var code = (uint) type;
        if (code < 0x40)
        {
            WriteByte((byte) (0x1 | (code << 2)));
            return;
        }

        WriteUInt16((ushort) (0x3 | (code << 2)));
    }

    /// <summary>
    /// A single stream object with its data.
    /// </summary>
    public void WriteSingle(RawStreamObject value)
    {
        WriteStart(value.Type, value.Data.Length);
        WriteBytes(value.Data.Span);
    }
}

/// <summary>
/// The sizes of encoded values, for the lengths of stream object headers.
/// </summary>
static class FsshttpbSize
{
    public static int BinaryItem(int length) => CompactUInt64.Size((ulong) length) + length;

    public static int StringItem(string value) => CompactUInt64.Size((ulong) value.Length) + value.Length * 2;

    public static int ExtendedGuidArray(IReadOnlyList<ExtendedGuid> values)
    {
        var size = CompactUInt64.Size((ulong) values.Count);
        foreach (var value in values)
        {
            size += value.Size;
        }

        return size;
    }

    public static int CellIdArray(IReadOnlyList<CellId> values)
    {
        var size = CompactUInt64.Size((ulong) values.Count);
        foreach (var value in values)
        {
            size += value.Size;
        }

        return size;
    }

    public static int StringItemArray(IReadOnlyList<string> values)
    {
        var size = CompactUInt64.Size((ulong) values.Count);
        foreach (var value in values)
        {
            size += StringItem(value);
        }

        return size;
    }
}

/// <summary>
/// A single stream object kept as its type and data, for the parts of the protocol the codec carries through without
/// interpreting.
/// </summary>
record RawStreamObject(StreamObjectType Type, ReadOnlyMemory<byte> Data);
