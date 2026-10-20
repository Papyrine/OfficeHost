/// <summary>
/// A node of the tree that holds a file stream ([MS-FSSHTTPD] 1.3.1, 2.2): an object of the cell, standing for the bytes
/// of the file under it, which follow those of the nodes before it.
/// </summary>
/// <param name="Children">The nodes under it, in the order of the file: its object references.</param>
abstract record FileNode(ExtendedGuid Id, ulong DataSize, IReadOnlyList<FileNode> Children)
{
    /// <summary>
    /// The data of the object of the node.
    /// </summary>
    public abstract ReadOnlyMemory<byte> ObjectData();

    /// <summary>
    /// The bytes of the file under the node, in order: the data of its Data Node Objects.
    /// </summary>
    public virtual IEnumerable<ReadOnlyMemory<byte>> Segments() => Children.SelectMany(_ => _.Segments());

    public byte[] ToArray()
    {
        if (DataSize > int.MaxValue)
        {
            throw new FsshttpdException($"The file of {DataSize} bytes is too large for an array");
        }

        var result = new byte[DataSize];
        var position = 0;
        foreach (var segment in Segments())
        {
            segment.Span.CopyTo(result.AsSpan(position));
            position += segment.Length;
        }

        return result;
    }
}

/// <summary>
/// An Intermediate or a Leaf Node Object: the signature of the file data under it (2.4), which its object data holds with
/// its size.
/// </summary>
abstract record NodeObject(ExtendedGuid Id, ReadOnlyMemory<byte> Signature, ulong DataSize, IReadOnlyList<FileNode> Children) :
    FileNode(Id, DataSize, Children);

/// <summary>
/// An Intermediate Node Object ([MS-FSSHTTPD] 2.2.2): the root of the tree, or a chunk with subchunks.
/// </summary>
record IntermediateNode(ExtendedGuid Id, ReadOnlyMemory<byte> Signature, ulong DataSize, IReadOnlyList<FileNode> Children) :
    NodeObject(Id, Signature, DataSize, Children)
{
    public override ReadOnlyMemory<byte> ObjectData() => new NodeObjectData(false, Signature, DataSize).ToArray();
}

/// <summary>
/// A Leaf Node Object ([MS-FSSHTTPD] 2.2.3): a chunk, whose child is the Data Node Object of its bytes, or whose children
/// are the Leaf Node Objects of its subchunks.
/// </summary>
record LeafNode(ExtendedGuid Id, ReadOnlyMemory<byte> Signature, ulong DataSize, IReadOnlyList<FileNode> Children, ReadOnlyMemory<byte>? DataHash = null) :
    NodeObject(Id, Signature, DataSize, Children)
{
    public override ReadOnlyMemory<byte> ObjectData() => new NodeObjectData(true, Signature, DataSize, DataHash).ToArray();
}

/// <summary>
/// A Data Node Object ([MS-FSSHTTPD] 2.2.4): the bytes of a chunk.
/// </summary>
record DataNode(ExtendedGuid Id, ReadOnlyMemory<byte> Data) :
    FileNode(Id, (ulong) Data.Length, [])
{
    public override ReadOnlyMemory<byte> ObjectData() => Data;

    public override IEnumerable<ReadOnlyMemory<byte>> Segments()
    {
        yield return Data;
    }
}
