/// <summary>
/// The data of an Intermediate Node Object ([MS-FSSHTTPD] 2.2.2.1) or a Leaf Node Object (2.2.3.1): the signature of the
/// file data under the node, its size, and for a leaf an optional data hash.
/// </summary>
/// <param name="DataHash">The Data Hash of a leaf (2.2.3.1), kept as it is: its algorithm is not fully specified (the
/// note &lt;3&gt;), so the nodes this host makes have none.</param>
record NodeObjectData(bool Leaf, ReadOnlyMemory<byte> Signature, ulong DataSize, ReadOnlyMemory<byte>? DataHash = null)
{
    StreamObjectType Type
    {
        get
        {
            if (Leaf)
            {
                return StreamObjectType.LeafNode;
            }

            return StreamObjectType.IntermediateNode;
        }
    }

    public static NodeObjectData Read(ReadOnlyMemory<byte> data)
    {
        var reader = new FsshttpbReader(data);
        bool leaf;
        if (reader.PeekStart(StreamObjectType.LeafNode))
        {
            leaf = true;
        }
        else if (reader.PeekStart(StreamObjectType.IntermediateNode))
        {
            leaf = false;
        }
        else
        {
            throw reader.Error("Not the start of a node object");
        }

        var type = leaf ? StreamObjectType.LeafNode : StreamObjectType.IntermediateNode;
        reader.ReadStart(type).EnsureEnd("a node object start");
        var signatureFields = reader.ReadStart(StreamObjectType.Signature);
        var signature = signatureFields.ReadBinaryItem();
        signatureFields.EnsureEnd("a node signature");
        var sizeFields = reader.ReadStart(StreamObjectType.DataSize);
        var size = sizeFields.ReadUInt64();
        sizeFields.EnsureEnd("a node data size");
        ReadOnlyMemory<byte>? hash = null;
        if (leaf && reader.PeekStart(StreamObjectType.DataHash))
        {
            var hashFields = reader.ReadStart(StreamObjectType.DataHash);
            hash = hashFields.ReadBinaryItem();
            hashFields.EnsureEnd("a node data hash");
        }

        reader.ReadEnd(type);
        reader.EnsureEnd("a node object");
        return new(leaf, signature, size, hash);
    }

    public byte[] ToArray()
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FsshttpbWriter(output);
        writer.WriteStart(Type, 0);
        writer.WriteStart(StreamObjectType.Signature, FsshttpbSize.BinaryItem(Signature.Length));
        writer.WriteBinaryItem(Signature.Span);
        writer.WriteStart(StreamObjectType.DataSize, 8);
        writer.WriteUInt64(DataSize);
        if (DataHash != null)
        {
            writer.WriteStart(StreamObjectType.DataHash, FsshttpbSize.BinaryItem(DataHash.Value.Length));
            writer.WriteBinaryItem(DataHash.Value.Span);
        }

        writer.WriteEnd(Type);
        return output.WrittenSpan.ToArray();
    }
}
