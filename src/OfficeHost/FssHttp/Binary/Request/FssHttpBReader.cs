/// <summary>
/// Reads FSSHTTPB ([MS-FSSHTTPB] 2.2.1) from memory without copying: binary data comes back as slices of the input.
/// Every read checks its bounds, so malformed or truncated input throws <see cref="FssHttpBException"/> and never reads
/// past the end, loops without consuming input, or allocates for a count the remaining bytes cannot hold.
/// </summary>
ref struct FssHttpBReader
{
    const int maxDepth = 32;

    ReadOnlyMemory<byte> memory;
    ReadOnlySpan<byte> span;
    int position;
    int offset;

    public FssHttpBReader(ReadOnlyMemory<byte> memory) :
        this(memory, 0)
    {
    }

    FssHttpBReader(ReadOnlyMemory<byte> memory, int offset)
    {
        this.memory = memory;
        span = memory.Span;
        this.offset = offset;
    }

    public int Remaining => span.Length - position;

    public bool AtEnd => position == span.Length;

    /// <summary>
    /// The position in the whole input, for the errors.
    /// </summary>
    public int Offset => offset + position;

    public FssHttpBException Error(string message) => new(message, Offset);

    ReadOnlySpan<byte> Take(int count)
    {
        if (count > Remaining)
        {
            throw Error($"{count} bytes expected, {Remaining} left");
        }

        var result = span.Slice(position, count);
        position += count;
        return result;
    }

    byte Peek()
    {
        if (AtEnd)
        {
            throw Error("Unexpected end of the data");
        }

        return span[position];
    }

    public byte ReadByte() => Take(1)[0];

    public ushort ReadUInt16() => BinaryPrimitives.ReadUInt16LittleEndian(Take(2));

    public uint ReadUInt32() => BinaryPrimitives.ReadUInt32LittleEndian(Take(4));

    public ulong ReadUInt64() => BinaryPrimitives.ReadUInt64LittleEndian(Take(8));

    public Guid ReadGuid() => new(Take(16));

    public ReadOnlyMemory<byte> ReadBytes(int count)
    {
        var start = position;
        Take(count);
        return memory.Slice(start, count);
    }

    public ReadOnlyMemory<byte> ReadRest() => ReadBytes(Remaining);

    /// <summary>
    /// A compact unsigned 64-bit integer (2.2.1.1): the lowest set bit of the first byte gives the size.
    /// </summary>
    public ulong ReadCompactUInt64()
    {
        var first = Peek();
        if (first == 0)
        {
            position++;
            return 0;
        }

        var size = BitOperations.TrailingZeroCount(first) + 1;
        if (size == 8)
        {
            // 2.2.1.1.9: the type byte 128, then the value
            position++;
            return ReadUInt64();
        }

        var bytes = Take(size);
        ulong value = 0;
        for (var index = size - 1; index >= 0; index--)
        {
            value = (value << 8) | bytes[index];
        }

        return value >> size;
    }

    /// <summary>
    /// An Extended GUID (2.2.1.7).
    /// </summary>
    public ExtendedGuid ReadExtendedGuid()
    {
        var at = Offset;
        var first = Peek();
        if (first == 0)
        {
            position++;
            return ExtendedGuid.Null;
        }

        uint value;
        if ((first & 0x07) == 0x04)
        {
            value = (uint) (ReadByte() >> 3);
        }
        else if ((first & 0x3F) == 0x20)
        {
            value = (uint) (ReadUInt16() >> 6);
        }
        else if ((first & 0x7F) == 0x40)
        {
            var bytes = Take(3);
            value = (uint) ((bytes[0] | (bytes[1] << 8) | (bytes[2] << 16)) >> 7);
        }
        else if (first == 0x80)
        {
            position++;
            value = ReadUInt32();
        }
        else
        {
            throw new FssHttpBException($"Invalid Extended GUID type 0x{first:X2}", at);
        }

        var guid = ReadGuid();
        if (guid == Guid.Empty)
        {
            throw new FssHttpBException("An Extended GUID that is not null has the empty GUID", at);
        }

        return new(guid, value);
    }

    /// <summary>
    /// A Serial Number (2.2.1.9).
    /// </summary>
    public SerialNumber ReadSerialNumber()
    {
        var at = Offset;
        var type = ReadByte();
        if (type == 0)
        {
            return SerialNumber.Null;
        }

        if (type != 0x80)
        {
            throw new FssHttpBException($"Invalid Serial Number type 0x{type:X2}", at);
        }

        var guid = ReadGuid();
        if (guid == Guid.Empty)
        {
            throw new FssHttpBException("A Serial Number that is not null has the empty GUID", at);
        }

        return new(guid, ReadUInt64());
    }

    /// <summary>
    /// A Cell ID (2.2.1.10).
    /// </summary>
    public CellId ReadCellId() => new(ReadExtendedGuid(), ReadExtendedGuid());

    /// <summary>
    /// A File Chunk Reference (2.2.1.2).
    /// </summary>
    public FileChunkReference ReadFileChunkReference() => new(ReadCompactUInt64(), ReadCompactUInt64());

    /// <summary>
    /// A count, checked against the bytes left for items of at least <paramref name="minimumSize"/> bytes.
    /// </summary>
    int ReadCount(int minimumSize)
    {
        var at = Offset;
        var count = ReadCompactUInt64();
        if (count > (ulong) (Remaining / minimumSize))
        {
            throw new FssHttpBException($"A count of {count} does not fit in the {Remaining} bytes left", at);
        }

        return (int) count;
    }

    /// <summary>
    /// An Extended GUID Array (2.2.1.8).
    /// </summary>
    public ExtendedGuid[] ReadExtendedGuidArray()
    {
        var result = new ExtendedGuid[ReadCount(1)];
        for (var index = 0; index < result.Length; index++)
        {
            result[index] = ReadExtendedGuid();
        }

        return result;
    }

    /// <summary>
    /// A Cell ID Array (2.2.1.11).
    /// </summary>
    public CellId[] ReadCellIdArray()
    {
        var result = new CellId[ReadCount(2)];
        for (var index = 0; index < result.Length; index++)
        {
            result[index] = ReadCellId();
        }

        return result;
    }

    /// <summary>
    /// A Binary Item (2.2.1.3), as a slice of the input.
    /// </summary>
    public ReadOnlyMemory<byte> ReadBinaryItem() => ReadBytes(ReadCount(1));

    /// <summary>
    /// A String Item (2.2.1.4): UTF-16 characters, not null-terminated.
    /// </summary>
    public string ReadStringItem() => Encoding.Unicode.GetString(Take(ReadCount(2) * 2));

    /// <summary>
    /// A String Item Array (2.2.1.14).
    /// </summary>
    public string[] ReadStringItemArray()
    {
        var result = new string[ReadCount(1)];
        for (var index = 0; index < result.Length; index++)
        {
            result[index] = ReadStringItem();
        }

        return result;
    }

    /// <summary>
    /// A stream object header (2.2.1.5): the two low bits of the first byte tell its kind.
    /// </summary>
    public StreamObjectHeader ReadHeader()
    {
        switch (Peek() & 0x3)
        {
            case 0:
            {
                // 2.2.1.5.1: 16-bit start
                var value = ReadUInt16();
                return new((StreamObjectType) ((value >> 3) & 0x3F), (value & 0x4) != 0, false, (ulong) (value >> 9));
            }
            case 1:
                // 2.2.1.5.3: 8-bit end
                return new((StreamObjectType) (ReadByte() >> 2), true, true, 0);
            case 2:
            {
                // 2.2.1.5.2: 32-bit start, with a Large Length when the length is 32767
                var value = ReadUInt32();
                var length = (ulong) (value >> 17);
                if (length == 32767)
                {
                    length = ReadCompactUInt64();
                }

                return new((StreamObjectType) ((value >> 3) & 0x3FFF), (value & 0x4) != 0, false, length);
            }
            default:
                // 2.2.1.5.4: 16-bit end
                return new((StreamObjectType) (ReadUInt16() >> 2), true, true, 0);
        }
    }

    public StreamObjectHeader PeekHeader()
    {
        var saved = position;
        var header = ReadHeader();
        position = saved;
        return header;
    }

    /// <summary>
    /// Whether the next stream object header is the start of an object of the type.
    /// </summary>
    public bool PeekStart(StreamObjectType type)
    {
        if (AtEnd)
        {
            return false;
        }

        var header = PeekHeader();
        return !header.End && header.Type == type;
    }

    /// <summary>
    /// Whether the next stream object header is an end, of any type.
    /// </summary>
    public bool PeekEnd()
    {
        if (AtEnd)
        {
            return false;
        }

        return PeekHeader().End;
    }

    /// <summary>
    /// Reads the start of a stream object of the type, and returns a reader of the data its length covers: the data of a
    /// single object, or the fields of a compound one that come before its first inner object.
    /// </summary>
    public FssHttpBReader ReadStart(StreamObjectType type)
    {
        var at = Offset;
        var header = ReadHeader();
        if (header.End || header.Type != type)
        {
            throw new FssHttpBException($"The start of {type} expected, {header} found", at);
        }

        if (header.Compound != type.IsCompound())
        {
            throw new FssHttpBException($"The compound bit of {type} is {header.Compound}", at);
        }

        return Slice(header, at);
    }

    FssHttpBReader Slice(StreamObjectHeader header, int at)
    {
        if (header.Length > (ulong) Remaining)
        {
            throw new FssHttpBException($"{header} does not fit in the {Remaining} bytes left", at);
        }

        var length = (int) header.Length;
        var start = position;
        position += length;
        return new(memory.Slice(start, length), offset + start);
    }

    /// <summary>
    /// Reads the end of a compound stream object of the type.
    /// </summary>
    public void ReadEnd(StreamObjectType type)
    {
        var at = Offset;
        var header = ReadHeader();
        if (!header.End || header.Type != type)
        {
            throw new FssHttpBException($"The end of {type} expected, {header} found", at);
        }
    }

    /// <summary>
    /// Reads a single stream object of any type, keeping its data as it is.
    /// </summary>
    public RawStreamObject ReadSingle()
    {
        var at = Offset;
        var header = ReadHeader();
        if (header.End || header.Compound || header.Type.IsCompound())
        {
            throw new FssHttpBException($"A single stream object expected, {header} found", at);
        }

        var data = Slice(header, at);
        return new(header.Type, data.ReadRest());
    }

    /// <summary>
    /// Skips whole stream objects up to the end header of the enclosing compound object, and returns their bytes.
    /// </summary>
    public ReadOnlyMemory<byte> ReadObjectsUntilEnd()
    {
        var start = position;
        while (!PeekEnd())
        {
            Skip(0);
        }

        return memory.Slice(start, position - start);
    }

    /// <summary>
    /// Skips a whole stream object, whose start the callers have peeked.
    /// </summary>
    void Skip(int depth)
    {
        if (depth > maxDepth)
        {
            throw Error("Stream objects nested too deeply");
        }

        var at = Offset;
        var header = ReadHeader();
        Slice(header, at);
        if (!header.Compound)
        {
            return;
        }

        while (!PeekEnd())
        {
            Skip(depth + 1);
        }

        ReadEnd(header.Type);
    }

    /// <summary>
    /// Fails if data is left: the end of a stream object or message must be where its length says.
    /// </summary>
    public void EnsureEnd(string what)
    {
        if (!AtEnd)
        {
            throw Error($"{Remaining} unexpected bytes at the end of {what}");
        }
    }
}
