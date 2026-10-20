/// <summary>
/// A chunk of a file ([MS-FSSHTTPD] 2.4): a range of the file and its signature, and the subchunks of a chunk larger than
/// 3 megabytes (2.4.1).
/// </summary>
record FileChunk(int Start, int Length, ReadOnlyMemory<byte> Signature, IReadOnlyList<FileChunk> SubChunks);

/// <summary>
/// How the signatures of zip items are made ([MS-FSSHTTPD] 2.4.1).
/// </summary>
/// <param name="XorItemSignatures">The signature of an item of at most 4096 bytes, header and data, is the exclusive or of
/// the signatures of its header and its data, as between clients and servers of version 2.2 or greater ([MS-FSSHTTP]
/// 2.2.5.10, 2.2.5.13), not their concatenation.</param>
/// <param name="CompressedDataCrc">The signature of the data of an item has the CRC-32 of its compressed data after the
/// CRC of the item, as the note &lt;6&gt; says Office does from the Office 2010 March 2015 CU. Word 16.0.20326 does not
/// (docs/fsshttp-notes.md).</param>
record ChunkingOptions(bool XorItemSignatures = true, bool CompressedDataCrc = false)
{
    /// <summary>
    /// The signatures of Word 16.0.20326 saving to SharePoint Server, and of this host.
    /// </summary>
    public static ChunkingOptions Default { get; } = new();

    /// <summary>
    /// The signatures of the example of [MS-FSSHTTPD] 3.1: concatenated, without the CRC of the compressed data.
    /// </summary>
    public static ChunkingOptions Original { get; } = new(false, false);
}

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
    /// chunk. No chunk at all when the file does not start with a local file header.
    /// </summary>
    static List<FileChunk> ZipChunks(ReadOnlySpan<byte> file, ChunkingOptions options)
    {
        var chunks = new List<FileChunk>();
        var position = 0;
        while (TryReadItem(file, position, out var item))
        {
            var dataStart = position + item.HeaderLength;
            if (item.CompressedSize > (ulong) (file.Length - dataStart))
            {
                // the data would extend past the end of the file
                break;
            }

            var dataLength = (int) item.CompressedSize;
            var header = file.Slice(position, item.HeaderLength);
            var data = file.Slice(dataStart, dataLength);
            var headerSignature = SHA1.HashData(header);
            var dataSignature = DataSignature(item, data, options);
            if (item.HeaderLength + dataLength <= smallItem)
            {
                chunks.Add(Chunk(file, position, item.HeaderLength + dataLength, Combine(headerSignature, dataSignature, options)));
            }
            else
            {
                chunks.Add(Chunk(file, position, item.HeaderLength, headerSignature));
                chunks.Add(Chunk(file, dataStart, dataLength, dataSignature));
            }

            position = dataStart + dataLength;
        }

        if (chunks.Count == 0 || position == file.Length)
        {
            return chunks;
        }

        var length = file.Length - position;
        byte[] signature;
        if (length <= megabyte)
        {
            signature = SHA1.HashData(file[position..]);
        }
        else
        {
            signature = UniqueSignature(file, position, length, 12);
        }

        chunks.Add(Chunk(file, position, length, signature));
        return chunks;
    }

    readonly record struct ZipItem(int HeaderLength, uint Crc, ulong CompressedSize, ulong UncompressedSize);

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

        item = new(headerLength, BinaryPrimitives.ReadUInt32LittleEndian(rest[14..]), compressed, uncompressed);
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
        for (var position = 0; position < file.Length; position += megabyte)
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

        var subChunks = new List<FileChunk>();
        for (var position = start; position < start + length; position += subChunkSize)
        {
            var subLength = Math.Min(subChunkSize, start + length - position);
            subChunks.Add(new(position, subLength, UniqueSignature(file, position, subLength, 8), []));
        }

        return new(start, length, signature, subChunks);
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
