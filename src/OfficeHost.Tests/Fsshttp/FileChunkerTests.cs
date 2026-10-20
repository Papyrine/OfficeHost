/// <summary>
/// The chunking of [MS-FSSHTTPD] 2.4: zip files at the boundaries of their items (2.4.1), any other file in chunks of
/// 1 megabyte (2.4.3).
/// </summary>
public class FileChunkerTests
{
    const int megabyte = 1024 * 1024;

    /// <summary>
    /// The zip of the example of [MS-FSSHTTPD] 3.1, from its Data Node Objects.
    /// </summary>
    static byte[] ExampleZip() => FileCell.Read(FsshttpdVectors.Storage()).Root.ToArray();

    static void AssertContiguous(IReadOnlyList<FileChunk> chunks, int length)
    {
        var position = 0;
        foreach (var chunk in chunks)
        {
            if (chunk.Start != position)
            {
                throw new($"The chunk at {chunk.Start} does not follow the one before, which ends at {position}");
            }

            var subPosition = chunk.Start;
            foreach (var subChunk in chunk.SubChunks)
            {
                if (subChunk.Start != subPosition)
                {
                    throw new($"The subchunk at {subChunk.Start} does not follow the one before, which ends at {subPosition}");
                }

                subPosition += subChunk.Length;
            }

            if (chunk.SubChunks.Count > 0 && subPosition != chunk.Start + chunk.Length)
            {
                throw new($"The subchunks of the chunk at {chunk.Start} do not cover it");
            }

            position += chunk.Length;
        }

        if (position != length)
        {
            throw new($"The chunks end at {position}, the file at {length}");
        }
    }

    /// <summary>
    /// The signatures of the example, byte for byte: each item is a chunk of at most 4096 bytes, whose signature is the
    /// SHA-1 of the local file header then the CRC and sizes of the data, and the central directory is a final chunk,
    /// signed by its SHA-1.
    /// </summary>
    [Test]
    public async Task Example_Signatures()
    {
        var zip = ExampleZip();
        await Assert.That(zip.Length).IsEqualTo(220);
        var chunks = FileChunker.Chunk(zip, ChunkingOptions.Original);
        await Assert.That(chunks.Select(_ => $"{_.Start}+{_.Length}").ToList()).IsEquivalentTo(["0+44", "44+44", "88+132"]);
        await Assert.That(Fsshttpb.Hex(chunks[0].Signature.Span))
            .IsEqualTo("F333D2A6BB6F43C9817AAB3A629D3C8A395F109D" + "8289D1F7" + "0500000000000000" + "0500000000000000");
        await Assert.That(Fsshttpb.Hex(chunks[1].Signature.Span))
            .IsEqualTo("912F5F635F88C7025ED9BD4896F41A62D3BCBEB4" + "473EB6FB" + "0500000000000000" + "0500000000000000");
        await Assert.That(Fsshttpb.Hex(chunks[2].Signature.Span)).IsEqualTo("49B53C0E99CA71E4D95371A66D006E60EA8FA6C6");
        await Assert.That(chunks.All(_ => _.SubChunks.Count == 0)).IsTrue();
    }

    static byte[] Xor(byte[] header, string data)
    {
        var result = Fsshttpb.Hex(data);
        for (var index = 0; index < header.Length; index++)
        {
            result[index] ^= header[index];
        }

        return result;
    }

    /// <summary>
    /// With the signatures of Word: a small item has the exclusive or of the SHA-1 of its header and of the CRC and sizes
    /// of its data, 20 bytes. Word's saves to SharePoint Server match this for every item (docs/fsshttp-notes.md).
    /// </summary>
    [Test]
    public async Task Example_DefaultSignatures()
    {
        var zip = ExampleZip();
        var chunks = FileChunker.Chunk(zip, ChunkingOptions.Default);
        var header = SHA1.HashData(zip.AsSpan(0, 39));
        await Assert.That(Fsshttpb.Hex(chunks[0].Signature.Span)).IsEqualTo(Fsshttpb.Hex(Xor(header, "8289D1F7" + "0500000000000000" + "0500000000000000")));
        await Assert.That(Fsshttpb.Hex(chunks[2].Signature.Span)).IsEqualTo("49B53C0E99CA71E4D95371A66D006E60EA8FA6C6");
    }

    /// <summary>
    /// With the CRC of the compressed data of the note &lt;6&gt;: the data signature is 24 bytes, and the exclusive or ends
    /// with the 4 bytes it has more.
    /// </summary>
    [Test]
    public async Task Example_CompressedDataCrc()
    {
        var zip = ExampleZip();
        var chunks = FileChunker.Chunk(zip, new(CompressedDataCrc: true));
        var header = SHA1.HashData(zip.AsSpan(0, 39));
        // the item is stored: its compressed data is "Hello", whose CRC is the CRC of the item
        await Assert.That(Fsshttpb.Hex(chunks[0].Signature.Span)).IsEqualTo(Fsshttpb.Hex(Xor(header, "8289D1F7" + "8289D1F7" + "0500000000000000" + "0500000000000000")));
    }

    /// <summary>
    /// An item larger than 4096 bytes is two chunks, its local file header and its data; the small items around it are a
    /// chunk each.
    /// </summary>
    [Test]
    public async Task Zip_LargeItem()
    {
        var large = Zips.Random(10_000, 1);
        byte[] zip =
        [
            ..StoredItem("small.txt", Zips.Text("small")),
            ..StoredItem("large.bin", large),
            ..StoredItem("last.txt", Zips.Text("last")),
            ..Zips.Text("central directory")
        ];
        var chunks = FileChunker.Chunk(zip, ChunkingOptions.Default);
        AssertContiguous(chunks, zip.Length);
        await Assert.That(chunks.Select(_ => _.Length).ToList()).IsEquivalentTo([30 + 9 + 5, 30 + 9, 10_000, 30 + 8 + 4, 17]);
        var headerStart = chunks[1].Start;
        await Assert.That(Fsshttpb.Hex(chunks[1].Signature.Span)).IsEqualTo(Fsshttpb.Hex(SHA1.HashData(zip.AsSpan(headerStart, 39))));
        var crc = Crc32.Compute(large);
        var dataSignature = new byte[20];
        BinaryPrimitives.WriteUInt32LittleEndian(dataSignature, crc);
        BinaryPrimitives.WriteUInt64LittleEndian(dataSignature.AsSpan(4), 10_000);
        BinaryPrimitives.WriteUInt64LittleEndian(dataSignature.AsSpan(12), 10_000);
        await Assert.That(Fsshttpb.Hex(chunks[2].Signature.Span)).IsEqualTo(Fsshttpb.Hex(dataSignature));
        var final = chunks[^1];
        await Assert.That(Fsshttpb.Hex(final.Signature.Span)).IsEqualTo(Fsshttpb.Hex(SHA1.HashData(zip.AsSpan(final.Start))));
    }

    /// <summary>
    /// A chunk larger than 3 megabytes has subchunks of 3 megabytes, the last one smaller, with unique signatures of 8
    /// bytes: the same at the same position for the same bytes, different elsewhere.
    /// </summary>
    [Test]
    public async Task Zip_SubChunks()
    {
        byte[] zip = [..StoredItem("zeros.bin", new byte[7 * megabyte]), ..Zips.Text("central directory")];
        var chunks = FileChunker.Chunk(zip, ChunkingOptions.Default);
        AssertContiguous(chunks, zip.Length);
        var data = chunks[1];
        await Assert.That(data.Length).IsEqualTo(7 * megabyte);
        await Assert.That(data.Signature.Length).IsEqualTo(20);
        await Assert.That(data.SubChunks.Select(_ => _.Length).ToList()).IsEquivalentTo([3 * megabyte, 3 * megabyte, megabyte]);
        await Assert.That(data.SubChunks.All(_ => _.Signature.Length == 8)).IsTrue();
        var signatures = data.SubChunks.Select(_ => Fsshttpb.Hex(_.Signature.Span)).ToList();
        await Assert.That(signatures.Distinct().Count()).IsEqualTo(3);
        var again = FileChunker.Chunk(zip, ChunkingOptions.Default)[1].SubChunks.Select(_ => Fsshttpb.Hex(_.Signature.Span)).ToList();
        await Assert.That(again).IsEquivalentTo(signatures);
    }

    static byte[] StoredItem(string name, byte[] data, byte[]? extra = null, uint? compressedSize = null)
    {
        extra ??= [];
        var header = new byte[30];
        BinaryPrimitives.WriteUInt32LittleEndian(header, 0x04034B50);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(4), 20);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(14), Crc32.Compute(data));
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(18), compressedSize ?? (uint) data.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(22), (uint) data.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), (ushort) name.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(28), (ushort) extra.Length);
        return [..header, ..Encoding.ASCII.GetBytes(name), ..extra, ..data];
    }

    /// <summary>
    /// The bytes after the last item, whatever they are, are the final chunk: signed by their SHA-1 up to 1 megabyte, by
    /// a unique signature of 12 bytes above, with subchunks above 3 megabytes.
    /// </summary>
    [Test]
    public async Task Zip_FinalChunk()
    {
        var item = StoredItem("a", Zips.Text("abc"));
        byte[] small = [..item, ..Zips.Text("not a header")];
        var chunks = FileChunker.Chunk(small, ChunkingOptions.Default);
        await Assert.That(chunks.Count).IsEqualTo(2);
        await Assert.That(Fsshttpb.Hex(chunks[1].Signature.Span)).IsEqualTo(Fsshttpb.Hex(SHA1.HashData(Zips.Text("not a header"))));

        byte[] large = [..item, ..Zips.Random(megabyte + 1, 2)];
        chunks = FileChunker.Chunk(large, ChunkingOptions.Default);
        await Assert.That(chunks[1].Length).IsEqualTo(megabyte + 1);
        await Assert.That(chunks[1].Signature.Length).IsEqualTo(12);
        await Assert.That(chunks[1].SubChunks.Count).IsEqualTo(0);

        byte[] huge = [..item, ..new byte[4 * megabyte]];
        chunks = FileChunker.Chunk(huge, ChunkingOptions.Default);
        AssertContiguous(chunks, huge.Length);
        await Assert.That(chunks[1].Signature.Length).IsEqualTo(12);
        await Assert.That(chunks[1].SubChunks.Select(_ => _.Length).ToList()).IsEquivalentTo([3 * megabyte, megabyte]);
    }

    /// <summary>
    /// The analysis stops at an item whose data would extend past the end of the file, and at a header that does not
    /// fit; when it stops before any item, the file is chunked as any other.
    /// </summary>
    [Test]
    public async Task Zip_Truncated()
    {
        var item = StoredItem("a", Zips.Text("abc"));
        var truncated = StoredItem("b", Zips.Text("def"), compressedSize: 1000);
        byte[] file = [..item, ..truncated];
        var chunks = FileChunker.Chunk(file, ChunkingOptions.Default);
        await Assert.That(chunks.Select(_ => _.Length).ToList()).IsEquivalentTo([item.Length, truncated.Length]);

        chunks = FileChunker.Chunk(truncated, ChunkingOptions.Default);
        await Assert.That(chunks.Count).IsEqualTo(1);
        await Assert.That(Fsshttpb.Hex(chunks[0].Signature.Span)).IsEqualTo(Fsshttpb.Hex(SHA1.HashData(truncated)));

        byte[] header = [0x50, 0x4B, 0x03, 0x04, 1, 2, 3];
        chunks = FileChunker.Chunk(header, ChunkingOptions.Default);
        await Assert.That(chunks.Single().Length).IsEqualTo(7);
        // the header of 31 bytes in 30
        chunks = FileChunker.Chunk(item.AsSpan(0, 30), ChunkingOptions.Default);
        await Assert.That(chunks.Single().Length).IsEqualTo(30);
    }

    /// <summary>
    /// The sizes of an item come from its Zip64 Extended Information Extra Field when it has one ([PKWARE-Zip] 4.5.3), and
    /// the other extra fields are skipped. A field that does not fit ends the extra fields.
    /// </summary>
    [Test]
    public async Task Zip_Zip64()
    {
        var data = Zips.Text("abc");
        var zip64 = new byte[20];
        BinaryPrimitives.WriteUInt16LittleEndian(zip64, 0x0001);
        BinaryPrimitives.WriteUInt16LittleEndian(zip64.AsSpan(2), 16);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64.AsSpan(4), 3);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64.AsSpan(12), 3);
        byte[] timestamp = [0x55, 0x54, 0x05, 0x00, 1, 2, 3, 4, 5];
        var item = StoredItem("a", data, [..timestamp, ..zip64], uint.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(item.AsSpan(22), uint.MaxValue);
        byte[] file = [..item, ..Zips.Text("END")];
        var chunks = FileChunker.Chunk(file, ChunkingOptions.Original);
        await Assert.That(chunks.Select(_ => _.Length).ToList()).IsEquivalentTo([item.Length, 3]);
        await Assert.That(Fsshttpb.Hex(chunks[0].Signature.Span[20..]))
            .IsEqualTo(Fsshttpb.Hex(BitConverter.GetBytes(Crc32.Compute(data))) + "0300000000000000" + "0300000000000000");

        // a field longer than the extra fields: the sizes stay at 0xFFFFFFFF, past the end of the file
        byte[] broken = [0x55, 0x54, 0x40, 0x00, 1];
        item = StoredItem("a", data, broken, uint.MaxValue);
        chunks = FileChunker.Chunk(item, ChunkingOptions.Original);
        await Assert.That(chunks.Single().Length).IsEqualTo(item.Length);
    }

    /// <summary>
    /// 2.4.3: any other file is chunked by megabyte, each signed by its SHA-1. An empty file has no chunk.
    /// </summary>
    [Test]
    public async Task Simple()
    {
        var file = Zips.Random(2 * megabyte + 100, 3);
        var chunks = FileChunker.Chunk(file, ChunkingOptions.Default);
        AssertContiguous(chunks, file.Length);
        await Assert.That(chunks.Select(_ => _.Length).ToList()).IsEquivalentTo([megabyte, megabyte, 100]);
        await Assert.That(Fsshttpb.Hex(chunks[1].Signature.Span)).IsEqualTo(Fsshttpb.Hex(SHA1.HashData(file.AsSpan(megabyte, megabyte))));
        await Assert.That(FileChunker.Chunk([], ChunkingOptions.Default).Count).IsEqualTo(0);
    }

    /// <summary>
    /// 2.4.3: in a file larger than 250 megabytes, each chunk has a unique signature of 12 bytes.
    /// </summary>
    [Test]
    public async Task Simple_LargeFile()
    {
        var file = new byte[250 * megabyte + 1];
        var chunks = FileChunker.Chunk(file, ChunkingOptions.Default);
        await Assert.That(chunks.Count).IsEqualTo(251);
        await Assert.That(chunks.All(_ => _.Signature.Length == 12)).IsTrue();
        await Assert.That(chunks.Select(_ => Fsshttpb.Hex(_.Signature.Span)).Distinct().Count()).IsEqualTo(251);
    }

    /// <summary>
    /// Zips made by <see cref="ZipArchive"/>, of random items, are chunked from end to end, every chunk at most 3
    /// megabytes or split into subchunks.
    /// </summary>
    [Test]
    public async Task Zip_Random()
    {
        var random = new Random(4);
        for (var iteration = 0; iteration < 20; iteration++)
        {
            var items = Enumerable.Range(0, random.Next(1, 30))
                .Select(_ => ($"item{_}.bin", Zips.Random(random.Next(0, 20_000), _)))
                .ToArray();
            var zip = Zips.Create(items);
            var chunks = FileChunker.Chunk(zip, ChunkingOptions.Default);
            AssertContiguous(chunks, zip.Length);
            await Assert.That(chunks.Count).IsGreaterThanOrEqualTo(items.Length + 1);
        }
    }

    [Test]
    [Arguments("", 0u)]
    [Arguments("Hello", 0xF7D18982u)]
    [Arguments("123456789", 0xCBF43926u)]
    public async Task Crc(string text, uint crc) =>
        await Assert.That(Crc32.Compute(Zips.Text(text))).IsEqualTo(crc);
}
