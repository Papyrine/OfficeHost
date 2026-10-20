/// <summary>
/// The tree of a file stream at a revision of its cell.
/// </summary>
record FileTree(ExtendedGuid RevisionId, FileNode Root);

/// <summary>
/// The data elements a write added, in the order of the example of [MS-FSSHTTPD] 3.1 (the Object Groups, then the
/// manifests, then the Storage Index), and the cell storage after it.
/// </summary>
record CellUpdate(IReadOnlyList<DataElement> Added, CellStorage Storage);

/// <summary>
/// A file stream in a cell storage, as [MS-FSSHTTPD] represents it: the Storage Manifest of its schema has the cell of
/// the file as its root (2.3), and the current revision of the cell has the root of a tree of node objects (2.2), whose
/// Data Node Objects hold the chunks of the file (2.4).
/// </summary>
static class FileCell
{
    static readonly Guid rootGuid = new("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");

    /// <summary>
    /// The Storage Manifest Schema GUID of a file (2.3).
    /// </summary>
    public static Guid SchemaId { get; } = new("0EB93394-571D-41E9-AAD3-880D92D31955");

    /// <summary>
    /// The root of the Storage Manifest, and of the Revision Manifest for the primary content stream (2.3).
    /// </summary>
    public static ExtendedGuid RootId { get; } = new(rootGuid, 2);

    /// <summary>
    /// The cell of the file (2.3).
    /// </summary>
    public static CellId CellId { get; } = new(new(rootGuid, 1), new(new("6F2A4665-42C8-46C7-BAB4-E28FDCE1E32B"), 1));

    const int maxDepth = 16;
    const int maxRevisions = 1000;

    /// <summary>
    /// The tree of the file at the current revision of its cell, whatever its chunking: the root may be a Leaf Node Object
    /// (the note &lt;1&gt;), a leaf may have subchunks, and the data may be in Object Data BLOBs or in Object Groups of earlier
    /// revisions.
    /// </summary>
    public static FileTree Read(CellStorage storage) => ReadRevision(storage).Tree;

    /// <summary>
    /// The tree of the file at the current revision, the Revision Manifest, and the objects of the revision.
    /// </summary>
    record Revision(FileTree Tree, RevisionManifest Manifest, Dictionary<ExtendedGuid, StoredObject> Objects);

    static Revision ReadRevision(CellStorage storage)
    {
        var index = storage.StorageIndex;
        var manifestMapping = index.Entries.OfType<StorageIndexManifestMapping>().FirstOrDefault() ??
                              throw new FsshttpdException("The Storage Index has no Storage Manifest mapping");
        var manifest = storage.Get<StorageManifest>(manifestMapping.MappingId);
        if (manifest.SchemaId != SchemaId)
        {
            throw new FsshttpdException($"The Storage Manifest has the schema {manifest.SchemaId}, not the one of a file");
        }

        var root = manifest.Roots.FirstOrDefault(_ => _.RootId == RootId) ??
                   throw new FsshttpdException($"The Storage Manifest has no root {RootId}");
        var cellMapping = index.Entries.OfType<StorageIndexCellMapping>().FirstOrDefault(_ => _.Cell == root.Cell) ??
                          throw new FsshttpdException($"The Storage Index has no mapping for the cell {root.Cell}");
        var cell = storage.Get<CellManifest>(cellMapping.MappingId);
        var revision = storage.GetRevision(cell.CurrentRevision);
        var rootObject = revision.Roots.FirstOrDefault(_ => _.RootId == RootId) ??
                         throw new FsshttpdException($"The revision {cell.CurrentRevision} has no root {RootId}");
        var objects = Objects(storage, revision);
        var reader = new TreeReader(storage, objects);
        var tree = new FileTree(cell.CurrentRevision, reader.Read(rootObject.ObjectId, false, 0));
        return new(tree, revision, objects);
    }

    /// <summary>
    /// An object of a revision, and the Object Group it is in.
    /// </summary>
    record StoredObject(IReadOnlyList<ExtendedGuid> References, ObjectGroupData Data, ExtendedGuid GroupId);

    /// <summary>
    /// The objects of a revision: those of its Object Groups, then those of its base revisions that it does not replace.
    /// </summary>
    static Dictionary<ExtendedGuid, StoredObject> Objects(CellStorage storage, RevisionManifest revision)
    {
        var objects = new Dictionary<ExtendedGuid, StoredObject>();
        var revisions = new HashSet<ExtendedGuid>();
        var current = revision;
        while (true)
        {
            if (!revisions.Add(current.RevisionId) || revisions.Count > maxRevisions)
            {
                throw new FsshttpdException($"The base revisions of {revision.RevisionId} loop or are more than {maxRevisions}");
            }

            foreach (var groupId in current.ObjectGroups)
            {
                var group = storage.Get<ObjectGroup>(groupId);
                if (group.Declarations.Count != group.Data.Count)
                {
                    throw new FsshttpdException($"The Object Group {groupId} has {group.Declarations.Count} declarations and {group.Data.Count} data");
                }

                for (var index = 0; index < group.Declarations.Count; index++)
                {
                    var data = group.Data[index];
                    objects.TryAdd(group.Declarations[index].ObjectId, new(data.ObjectReferences, data, groupId));
                }
            }

            if (current.BaseRevisionId.IsNull)
            {
                return objects;
            }

            current = storage.GetRevision(current.BaseRevisionId);
        }
    }

    class TreeReader(CellStorage storage, Dictionary<ExtendedGuid, StoredObject> objects)
    {
        HashSet<ExtendedGuid> visited = [];

        ReadOnlyMemory<byte> Data(ExtendedGuid id, ObjectGroupData data) =>
            data switch
            {
                ObjectData value => value.Data,
                ObjectDataBlobReference value => storage.Get<ObjectDataBlob>(value.BlobId).Data,
                _ => throw new FsshttpdException($"The data of the object {id} is excluded")
            };

        /// <summary>
        /// A node and the nodes under it. Under a leaf, an object without references is a Data Node Object; any other
        /// object is a node object. Each object is in the tree once, so that a tree cannot loop or stand for more data
        /// than its objects hold.
        /// </summary>
        public FileNode Read(ExtendedGuid id, bool underLeaf, int depth)
        {
            if (depth > maxDepth)
            {
                throw new FsshttpdException($"The tree of node objects is deeper than {maxDepth}");
            }

            if (!visited.Add(id))
            {
                throw new FsshttpdException($"The object {id} is in the tree more than once");
            }

            if (!objects.TryGetValue(id, out var stored))
            {
                throw new FsshttpdException($"The revision has no object {id}");
            }

            var data = Data(id, stored.Data);
            if (underLeaf && stored.References.Count == 0)
            {
                return new DataNode(id, data);
            }

            NodeObjectData node;
            try
            {
                node = NodeObjectData.Read(data);
            }
            catch (FsshttpbException exception)
            {
                throw new FsshttpdException($"The object {id} is not a node object: {exception.Message}", exception);
            }

            var children = new List<FileNode>();
            ulong size = 0;
            foreach (var reference in stored.References)
            {
                var child = Read(reference, node.Leaf, depth + 1);
                if (child.DataSize > node.DataSize - size)
                {
                    throw new FsshttpdException($"The children of the node {id} have more than its {node.DataSize} bytes");
                }

                size += child.DataSize;
                children.Add(child);
            }

            if (size != node.DataSize)
            {
                throw new FsshttpdException($"The node {id} has {node.DataSize} bytes, its children {size}");
            }

            if (!node.Leaf)
            {
                return new IntermediateNode(id, node.Signature, node.DataSize, children);
            }

            if (children is not [DataNode] && !children.All(_ => _ is LeafNode))
            {
                throw new FsshttpdException($"The leaf {id} has neither a single Data Node Object nor Leaf Node Objects");
            }

            return new LeafNode(id, node.Signature, node.DataSize, children, node.DataHash);
        }
    }

    /// <summary>
    /// Writes a file as the next revision of its cell storage, or of a new one. The chunks of the previous revision that
    /// have the same signature, size and bytes are kept with their objects, and the new revision lists their Object
    /// Groups again, as Word does, with no base revision. It has new Object Groups only for the new objects, one each,
    /// as in the example of 3.1. The other roots of the previous revision, streams opaque to [MS-FSSHTTPD] (2.3), are kept
    /// the same way. The Storage Index, Storage Manifest and Cell Manifest keep their identifiers. The storage after the
    /// write has only the data elements reachable from its Storage Index. When the file has not changed, nothing is added.
    /// The data elements hold slices of <paramref name="file"/>, which must not change after.
    /// </summary>
    public static CellUpdate Write(ReadOnlyMemory<byte> file, CellStorage? previous, IdAllocator ids, ChunkingOptions options)
    {
        Revision? current = null;
        var reusable = new Dictionary<string, FileNode>();
        if (previous != null)
        {
            current = ReadRevision(previous);
            if (SameBytes(current.Tree.Root, file.Span))
            {
                return new([], previous);
            }

            if (current.Tree.Root is IntermediateNode previousRoot)
            {
                foreach (var child in previousRoot.Children.Cast<NodeObject>())
                {
                    if (child.Signature.Length > 0)
                    {
                        reusable.TryAdd(Key(child.Signature, child.DataSize), child);
                    }
                }
            }
        }

        var rootId = ids.NewId();
        var children = new List<FileNode>();
        var reused = new HashSet<FileNode>(ReferenceEqualityComparer.Instance);
        foreach (var chunk in FileChunker.Chunk(file.Span, options))
        {
            var bytes = file.Slice(chunk.Start, chunk.Length);
            if (reusable.TryGetValue(Key(chunk.Signature, (ulong) chunk.Length), out var existing) &&
                !reused.Contains(existing) &&
                SameBytes(existing, bytes.Span))
            {
                reused.Add(existing);
                children.Add(existing);
                continue;
            }

            children.Add(Node(chunk, file, ids));
        }

        var root = new IntermediateNode(rootId, ReadOnlyMemory<byte>.Empty, (ulong) file.Length, children);
        var groups = new List<DataElement>();
        var listed = new GroupList(current?.Objects ?? []);
        var queue = new Queue<FileNode>([root]);
        while (queue.TryDequeue(out var node))
        {
            var group = Group(node, ids);
            groups.Add(group);
            listed.Add(group.Id);
            foreach (var child in node.Children)
            {
                if (reused.Contains(child))
                {
                    listed.AddObjects(child.Id);
                }
                else
                {
                    queue.Enqueue(child);
                }
            }
        }

        var roots = new List<RevisionManifestRoot>
        {
            new(RootId, root.Id)
        };
        foreach (var other in current?.Manifest.Roots ?? [])
        {
            if (other.RootId != RootId)
            {
                roots.Add(other);
                listed.AddObjects(other.ObjectId);
            }
        }

        var revisionId = ids.NewId();
        var revision = new RevisionManifest(ids.NewId(), ids.NewSerialNumber(), revisionId, ExtendedGuid.Null, roots, listed.Ids);
        var added = new List<DataElement>(groups);

        var previousIndex = previous?.StorageIndex;
        var manifestMapping = previousIndex?.Entries.OfType<StorageIndexManifestMapping>().FirstOrDefault();
        if (manifestMapping == null)
        {
            var manifest = new StorageManifest(ids.NewId(), ids.NewSerialNumber(), SchemaId, [new(RootId, CellId)]);
            added.Add(manifest);
            manifestMapping = new(manifest.Id, ids.NewSerialNumber());
        }

        var cellMapping = previousIndex?.Entries.OfType<StorageIndexCellMapping>().FirstOrDefault(_ => _.Cell == CellId);
        ExtendedGuid cellManifestId;
        if (cellMapping == null)
        {
            cellManifestId = ids.NewId();
            cellMapping = new(CellId, cellManifestId, ids.NewSerialNumber());
        }
        else
        {
            cellManifestId = cellMapping.MappingId;
        }

        added.Add(new CellManifest(cellManifestId, ids.NewSerialNumber(), revisionId));
        added.Add(revision);

        var entries = new List<StorageIndexEntry>
        {
            manifestMapping,
            cellMapping,
            new StorageIndexRevisionMapping(revisionId, revision.Id, ids.NewSerialNumber())
        };
        var storageIndexId = previous?.StorageIndexId ?? ids.NewId();
        added.Add(new StorageIndex(storageIndexId, ids.NewSerialNumber(), entries));
        var storage = (previous ?? CellStorage.Create([], storageIndexId)).With(added, storageIndexId).Prune();
        return new(added, storage);
    }

    /// <summary>
    /// The Object Groups of a revision, in order and each once: the new ones, and those of the objects it keeps from the
    /// previous revision, with the objects under them.
    /// </summary>
    class GroupList(Dictionary<ExtendedGuid, StoredObject> objects)
    {
        HashSet<ExtendedGuid> groups = [];
        HashSet<ExtendedGuid> visited = [];

        public List<ExtendedGuid> Ids { get; } = [];

        public void Add(ExtendedGuid groupId)
        {
            if (groups.Add(groupId))
            {
                Ids.Add(groupId);
            }
        }

        public void AddObjects(ExtendedGuid objectId)
        {
            var pending = new Stack<ExtendedGuid>([objectId]);
            while (pending.TryPop(out var id))
            {
                if (!visited.Add(id))
                {
                    continue;
                }

                if (!objects.TryGetValue(id, out var stored))
                {
                    throw new FsshttpdException($"The previous revision has no object {id}");
                }

                Add(stored.GroupId);
                foreach (var reference in stored.References)
                {
                    pending.Push(reference);
                }
            }
        }
    }

    static string Key(ReadOnlyMemory<byte> signature, ulong size) => $"{Convert.ToHexString(signature.Span)}:{size}";

    static bool SameBytes(FileNode node, ReadOnlySpan<byte> bytes)
    {
        if (node.DataSize != (ulong) bytes.Length)
        {
            return false;
        }

        var position = 0;
        foreach (var segment in node.Segments())
        {
            if (!segment.Span.SequenceEqual(bytes.Slice(position, segment.Length)))
            {
                return false;
            }

            position += segment.Length;
        }

        return true;
    }

    /// <summary>
    /// The node of a chunk: a Leaf Node Object with the Data Node Object of its bytes, or for a chunk with subchunks an
    /// Intermediate Node Object with a Leaf Node Object for each (2.4.1).
    /// </summary>
    static FileNode Node(FileChunk chunk, ReadOnlyMemory<byte> file, IdAllocator ids)
    {
        var id = ids.NewId();
        if (chunk.SubChunks.Count == 0)
        {
            return new LeafNode(id, chunk.Signature, (ulong) chunk.Length, [new DataNode(ids.NewId(), file.Slice(chunk.Start, chunk.Length))]);
        }

        var subChunks = chunk.SubChunks.Select(_ => Node(_, file, ids)).ToList();
        return new IntermediateNode(id, chunk.Signature, (ulong) chunk.Length, subChunks);
    }

    /// <summary>
    /// The Object Group of a node object, alone in it (2.2.1): partition 1, the children as object references, no cell
    /// references.
    /// </summary>
    static ObjectGroup Group(FileNode node, IdAllocator ids)
    {
        var data = node.ObjectData();
        var references = node.Children.Select(_ => _.Id).ToList();
        return new(
            ids.NewId(),
            ids.NewSerialNumber(),
            null,
            [new ObjectDeclaration(node.Id, 1, (ulong) data.Length, (ulong) references.Count, 0)],
            null,
            [new ObjectData(references, [], data)]);
    }
}
