/// <summary>
/// Splits a file into chunks ([MS-FSSHTTPD] 2.4): a zip file at the boundaries of its items (2.4.1), and any other file in
/// chunks of 1 megabyte (2.4.3). The RDC analysis of 2.4.2 is not done: it only makes other chunks, and the tree of
/// node objects of any chunking is read the same.
/// </summary>
static class FileChunker
{
    const int megabyte = 1024 * 1024;
    const int smallItem = 4096;
    const int subChunkSize = 3 * megabyte;
    const long largeFile = 250L * megabyte;
    const uint localFileHeaderSignature = 0x04034B50;
    const int localFileHeaderSize = 30;
    const ushort zip64ExtraField = 0x0001;

    public static IReadOnlyList<FileChunk> Chunk(ReadOnlySpan<byte> file, ChunkingOptions options)
    {
        var chunks = ZipChunks(file, options);
        if (chunks.Count > 0)
        {
            return chunks;
        }

        return SimpleChunks(file);
    }

    /// <summary>
    /// 2.4.1: from the start of the file, each local file header is a chunk and the data that follows it another, or both
    /// one chunk when they take at most 4096 bytes. The bytes after the last item, the central directory, are a final
    /// chunk. No chunk at all when the file does not start with a local file header. An item written with a data
    /// descriptor ([PKWARE-Zip] 4.3.9, as a streaming zip writer writes, its local header without CRC or sizes) has
    /// them from the central directory, and its data chunk ends after the descriptor.
    /// </summary>
    static List<FileChunk> ZipChunks(ReadOnlySpan<byte> file, ChunkingOptions options)
    {
        var chunks = new List<FileChunk>();
        var position = 0;
        Dictionary<long, ZipItem>? central = null;
        while (TryReadItem(file, position, out var item))
        {
            var dataStart = position + item.HeaderLength;
            if (item.HasDataDescriptor)
            {
                central ??= CentralDirectory(file);
                if (!central.TryGetValue(position, out var sizes))
                {
                    break;
                }

                item = item with
                {
                    Crc = sizes.Crc,
                    CompressedSize = sizes.CompressedSize,
                    UncompressedSize = sizes.UncompressedSize
                };
            }

            if (item.CompressedSize > (ulong) (file.Length - dataStart))
            {
                // the data would extend past the end of the file
                break;
            }

            var dataLength = (int) item.CompressedSize;
            var descriptor = 0;
            if (item.HasDataDescriptor &&
                (descriptor = DescriptorLength(file, dataStart + dataLength)) == 0)
            {
                break;
            }

            var header = file.Slice(position, item.HeaderLength);
            var data = file.Slice(dataStart, dataLength);
            var headerSignature = SHA1.HashData(header);
            var dataSignature = DataSignature(item, data, options);
            if (item.HeaderLength + dataLength + descriptor <= smallItem)
            {
                chunks.Add(Chunk(file, position, item.HeaderLength + dataLength + descriptor, Combine(headerSignature, dataSignature, options)));
            }
            else
            {
                chunks.Add(Chunk(file, position, item.HeaderLength, headerSignature));
                chunks.Add(Chunk(file, dataStart, dataLength + descriptor, dataSignature));
            }

            position = dataStart + dataLength + descriptor;
        }

        if (chunks.Count == 0 || position == file.Length)
        {
            return chunks;
        }

        var length = file.Length - position;
        if (length <= megabyte)
        {
            chunks.Add(Chunk(file, position, length, SHA1.HashData(file[position..])));
            return chunks;
        }

        if (length <= subChunkSize)
        {
            chunks.Add(Chunk(file, position, length, UniqueSignature(file, position, length, 12)));
            return chunks;
        }

        // with subchunks, the chunk is signed by theirs, unique as they are: its bytes are hashed once, not twice
        var subChunks = SubChunks(file, position, length);
        chunks.Add(new(position, length, UniqueSignature(position, subChunks), subChunks));
        return chunks;
    }

    readonly record struct ZipItem(int HeaderLength, uint Crc, ulong CompressedSize, ulong UncompressedSize, bool HasDataDescriptor = false);

    const uint dataDescriptorSignature = 0x08074B50;
    const uint centralFileHeaderSignature = 0x02014B50;
    const uint endOfCentralDirectorySignature = 0x06054B50;
    const uint zip64EndOfCentralDirectorySignature = 0x06064B50;
    const uint zip64LocatorSignature = 0x07064B50;

    /// <summary>
    /// The length of the data descriptor at a position (4.3.9): with or without its signature, with sizes of 4 or 8
    /// bytes, the one followed by a local or a central file header. 0 when none is.
    /// </summary>
    static int DescriptorLength(ReadOnlySpan<byte> file, int position)
    {
        var signed = file.Length - position >= 4 &&
                     BinaryPrimitives.ReadUInt32LittleEndian(file[position..]) == dataDescriptorSignature;
        int[] lengths = signed ? [16, 24] : [12, 20];
        foreach (var length in lengths)
        {
            var next = position + length;
            if (file.Length - next >= 4 &&
                BinaryPrimitives.ReadUInt32LittleEndian(file[next..]) is localFileHeaderSignature or centralFileHeaderSignature)
            {
                return length;
            }
        }

        return 0;
    }

    /// <summary>
    /// The CRC and sizes of the items of the central directory (4.3.12), by the position of their local header, from the
    /// end of central directory record (4.3.16), or its Zip64 record (4.3.14); none when the file has no central
    /// directory that can be read.
    /// </summary>
    static Dictionary<long, ZipItem> CentralDirectory(ReadOnlySpan<byte> file)
    {
        var items = new Dictionary<long, ZipItem>();
        if (!TryFindCentralDirectory(file, out var offset, out var count))
        {
            return items;
        }

        var position = offset;
        for (long index = 0; index < count; index++)
        {
            if (position < 0 ||
                file.Length - position < 46 ||
                BinaryPrimitives.ReadUInt32LittleEndian(file[(int) position..]) != centralFileHeaderSignature)
            {
                break;
            }

            var entry = file[(int) position..];
            ulong compressed = BinaryPrimitives.ReadUInt32LittleEndian(entry[20..]);
            ulong uncompressed = BinaryPrimitives.ReadUInt32LittleEndian(entry[24..]);
            long localHeader = BinaryPrimitives.ReadUInt32LittleEndian(entry[42..]);
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(entry[28..]);
            var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(entry[30..]);
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(entry[32..]);
            var length = 46 + nameLength + extraLength + commentLength;
            if (length > entry.Length)
            {
                break;
            }

            // the Zip64 field has, in order, the values the header has at their maximum (4.5.3)
            var extra = entry.Slice(46 + nameLength, extraLength);
            while (extra.Length >= 4)
            {
                var id = BinaryPrimitives.ReadUInt16LittleEndian(extra);
                var size = BinaryPrimitives.ReadUInt16LittleEndian(extra[2..]);
                if (size > extra.Length - 4)
                {
                    break;
                }

                var field = extra.Slice(4, size);
                if (id == zip64ExtraField)
                {
                    if (uncompressed == uint.MaxValue && field.Length >= 8)
                    {
                        uncompressed = BinaryPrimitives.ReadUInt64LittleEndian(field);
                        field = field[8..];
                    }

                    if (compressed == uint.MaxValue && field.Length >= 8)
                    {
                        compressed = BinaryPrimitives.ReadUInt64LittleEndian(field);
                        field = field[8..];
                    }

                    if (localHeader == uint.MaxValue && field.Length >= 8)
                    {
                        localHeader = (long) BinaryPrimitives.ReadUInt64LittleEndian(field);
                    }
                }

                extra = extra[(4 + size)..];
            }

            items.TryAdd(localHeader, new(0, BinaryPrimitives.ReadUInt32LittleEndian(entry[16..]), compressed, uncompressed));
            position += length;
        }

        return items;
    }

    static bool TryFindCentralDirectory(ReadOnlySpan<byte> file, out long offset, out long count)
    {
        offset = 0;
        count = 0;
        // the record is 22 bytes and a comment of up to 65535
        for (var position = file.Length - 22; position >= Math.Max(0, file.Length - 22 - ushort.MaxValue); position--)
        {
            if (BinaryPrimitives.ReadUInt32LittleEndian(file[position..]) != endOfCentralDirectorySignature)
            {
                continue;
            }

            var record = file[position..];
            count = BinaryPrimitives.ReadUInt16LittleEndian(record[10..]);
            offset = BinaryPrimitives.ReadUInt32LittleEndian(record[16..]);
            if ((count == ushort.MaxValue || offset == uint.MaxValue) &&
                position >= 20 &&
                BinaryPrimitives.ReadUInt32LittleEndian(file[(position - 20)..]) == zip64LocatorSignature)
            {
                var zip64 = (long) BinaryPrimitives.ReadUInt64LittleEndian(file[(position - 12)..]);
                if (zip64 < 0 ||
                    file.Length - zip64 < 56 ||
                    BinaryPrimitives.ReadUInt32LittleEndian(file[(int) zip64..]) != zip64EndOfCentralDirectorySignature)
                {
                    return false;
                }

                count = (long) BinaryPrimitives.ReadUInt64LittleEndian(file[((int) zip64 + 32)..]);
                offset = (long) BinaryPrimitives.ReadUInt64LittleEndian(file[((int) zip64 + 48)..]);
            }

            return true;
        }

        return false;
    }

    /// <summary>
    /// A local file header ([PKWARE-Zip] 4.3.7) that fits in the file, with the sizes of its Zip64 Extended Information
    /// Extra Field (4.5.3) when it has one.
    /// </summary>
    static bool TryReadItem(ReadOnlySpan<byte> file, int position, out ZipItem item)
    {
        item = default;
        var rest = file[position..];
        if (rest.Length < localFileHeaderSize ||
            BinaryPrimitives.ReadUInt32LittleEndian(rest) != localFileHeaderSignature)
        {
            return false;
        }

        var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(rest[26..]);
        var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(rest[28..]);
        var headerLength = localFileHeaderSize + nameLength + extraLength;
        if (headerLength > rest.Length)
        {
            return false;
        }

        ulong compressed = BinaryPrimitives.ReadUInt32LittleEndian(rest[18..]);
        ulong uncompressed = BinaryPrimitives.ReadUInt32LittleEndian(rest[22..]);
        var extra = rest.Slice(localFileHeaderSize + nameLength, extraLength);
        while (extra.Length >= 4)
        {
            var id = BinaryPrimitives.ReadUInt16LittleEndian(extra);
            var size = BinaryPrimitives.ReadUInt16LittleEndian(extra[2..]);
            var fieldData = extra[4..];
            if (size > fieldData.Length)
            {
                break;
            }

            // in a local file header, the field has both sizes: the uncompressed, then the compressed
            if (id == zip64ExtraField && size >= 16)
            {
                uncompressed = BinaryPrimitives.ReadUInt64LittleEndian(fieldData);
                compressed = BinaryPrimitives.ReadUInt64LittleEndian(fieldData[8..]);
            }

            extra = fieldData[size..];
        }

        // bit 3 of the general purpose flags: the CRC and sizes follow the data (4.4.4)
        var hasDataDescriptor = (BinaryPrimitives.ReadUInt16LittleEndian(rest[6..]) & 0x08) != 0;
        item = new(headerLength, BinaryPrimitives.ReadUInt32LittleEndian(rest[14..]), compressed, uncompressed, hasDataDescriptor);
        return true;
    }

    /// <summary>
    /// The signature of the data of an item: its CRC, the CRC of its compressed data with
    /// <see cref="ChunkingOptions.CompressedDataCrc"/>, its compressed size and its uncompressed size.
    /// </summary>
    static byte[] DataSignature(ZipItem item, ReadOnlySpan<byte> data, ChunkingOptions options)
    {
        var crcs = 4;
        if (options.CompressedDataCrc)
        {
            crcs = 8;
        }

        var signature = new byte[crcs + 16];
        BinaryPrimitives.WriteUInt32LittleEndian(signature, item.Crc);
        if (options.CompressedDataCrc)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(signature.AsSpan(4), Crc32.Compute(data));
        }

        BinaryPrimitives.WriteUInt64LittleEndian(signature.AsSpan(crcs), item.CompressedSize);
        BinaryPrimitives.WriteUInt64LittleEndian(signature.AsSpan(crcs + 8), item.UncompressedSize);
        return signature;
    }

    /// <summary>
    /// The signature of a small item: the two signatures concatenated, or their exclusive or followed by the extra bytes of
    /// the longer one, which is the data signature (20 or 24 bytes, against the 20 of the SHA-1 of the header).
    /// </summary>
    static byte[] Combine(byte[] header, byte[] data, ChunkingOptions options)
    {
        if (!options.XorItemSignatures)
        {
            return [..header, ..data];
        }

        var result = data.ToArray();
        for (var index = 0; index < header.Length; index++)
        {
            result[index] ^= header[index];
        }

        return result;
    }

    /// <summary>
    /// 2.4.3: chunks of 1 megabyte, signed by their SHA-1, or by a unique signature of 12 bytes in a file larger than 250
    /// megabytes.
    /// </summary>
    static List<FileChunk> SimpleChunks(ReadOnlySpan<byte> file)
    {
        var chunks = new List<FileChunk>();
        for (var position = 0; position < file.Length; position += Math.Min(megabyte, file.Length - position))
        {
            var length = Math.Min(megabyte, file.Length - position);
            byte[] signature;
            if (file.Length <= largeFile)
            {
                signature = SHA1.HashData(file.Slice(position, length));
            }
            else
            {
                signature = UniqueSignature(file, position, length, 12);
            }

            chunks.Add(new(position, length, signature, []));
        }

        return chunks;
    }

    /// <summary>
    /// A chunk, with subchunks of 3 megabytes when it is larger than that, each with a unique signature of 8 bytes.
    /// </summary>
    static FileChunk Chunk(ReadOnlySpan<byte> file, int start, int length, byte[] signature)
    {
        if (length <= subChunkSize)
        {
            return new(start, length, signature, []);
        }

        return new(start, length, signature, SubChunks(file, start, length));
    }

    /// <summary>
    /// The subchunks of 3 megabytes of a chunk, each with a unique signature of 8 bytes. Positions step by the length of
    /// each, which ends at the end of the chunk: a step of 3 megabytes past it would overflow near 2 GB.
    /// </summary>
    static List<FileChunk> SubChunks(ReadOnlySpan<byte> file, int start, int length)
    {
        var subChunks = new List<FileChunk>();
        var end = start + length;
        for (var position = start; position < end;)
        {
            var subLength = Math.Min(subChunkSize, end - position);
            subChunks.Add(new(position, subLength, UniqueSignature(file, position, subLength, 8), []));
            position += subLength;
        }

        return subChunks;
    }

    /// <summary>
    /// The unique signature of 12 bytes of a chunk with subchunks: of its position and their signatures.
    /// </summary>
    static byte[] UniqueSignature(int start, List<FileChunk> subChunks)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA1);
        Span<byte> position = stackalloc byte[8];
        BinaryPrimitives.WriteInt64LittleEndian(position, start);
        hash.AppendData(position);
        foreach (var subChunk in subChunks)
        {
            hash.AppendData(subChunk.Signature.Span);
        }

        return hash.GetHashAndReset()[..12];
    }

    /// <summary>
    /// The signatures that [MS-FSSHTTPD] only requires to be unique, which Office makes random (the notes &lt;7&gt; and
    /// &lt;9&gt;): here the start of the SHA-1 of the position and the bytes of the chunk, so that the same bytes at the same
    /// position get the same signature, and the chunk can be reused by the next revision.
    /// </summary>
    static byte[] UniqueSignature(ReadOnlySpan<byte> file, int start, int length, int size)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA1);
        Span<byte> position = stackalloc byte[8];
        BinaryPrimitives.WriteInt64LittleEndian(position, start);
        hash.AppendData(position);
        hash.AppendData(file.Slice(start, length));
        return hash.GetHashAndReset()[..size];
    }
}
