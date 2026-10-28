/// <summary>
/// A file stream in a cell storage ([MS-FSSHTTPD]): the tree of node objects read from data elements, and written from
/// the bytes of a file.
/// </summary>
public class FileCellTests
{
    const int megabyte = 1024 * 1024;

    static IdAllocator Ids() => new(FssHttpB.Guid("5F9E42C5-0A57-4C6E-9D3A-6B1F7C2E8D41"));

    /// <summary>
    /// The same bytes in the same order, in linear time: <c>IsEquivalentTo</c> ignores the order, at a quadratic cost for
    /// files of megabytes.
    /// </summary>
    static async Task AssertBytes(byte[] actual, byte[] expected)
    {
        await Assert.That(actual.Length).IsEqualTo(expected.Length);
        await Assert.That(actual.AsSpan().SequenceEqual(expected)).IsTrue();
    }

    /// <summary>
    /// The description of data elements, with each Extended GUID and Serial Number replaced by its order of appearance:
    /// the same structure with other identifiers gives the same text.
    /// </summary>
    static string Canonical(IEnumerable<DataElement> elements)
    {
        var ids = new Dictionary<string, int>();
        return Regex.Replace(
            FssHttpB.Describe(elements.ToList()),
            @"\{[0-9a-f-]{36}\},\d+",
            _ =>
            {
                if (!ids.TryGetValue(_.Value, out var number))
                {
                    number = ids.Count;
                    ids[_.Value] = number;
                }

                return $"#{number}";
            });
    }

    /// <summary>
    /// The example of [MS-FSSHTTPD] 3.1 is a tree of an Intermediate Node Object with no signature, over a Leaf Node
    /// Object for each chunk, each over the Data Node Object of its bytes, which make the zip.
    /// </summary>
    [Test]
    public async Task Example_Read()
    {
        var tree = FileCell.Read(FssHttpDVectors.Storage());
        var root = (IntermediateNode) tree.Root;
        await Assert.That(root.Signature.Length).IsEqualTo(0);
        await Assert.That(root.DataSize).IsEqualTo(220ul);
        await Assert.That(root.Children.Select(_ => _.DataSize).ToList()).IsEquivalentTo([44ul, 44ul, 132ul]);
        foreach (var child in root.Children)
        {
            var leaf = (LeafNode) child;
            await Assert.That(leaf.Children.Single()).IsTypeOf<DataNode>();
            await Assert.That(leaf.DataHash).IsNull();
        }

        var zip = root.ToArray();
        await Assert.That(Zips.Read(zip)).IsEquivalentTo(new Dictionary<string, string>
        {
            ["Hello.txt"] = "Hello",
            ["World.txt"] = "World"
        });
    }

    /// <summary>
    /// The zip of the example, written with the signatures of the example, makes the data elements of the example in the
    /// same order and with the same content, but for the identifiers and Serial Numbers.
    /// </summary>
    [Test]
    public async Task Example_Write()
    {
        var storage = FssHttpDVectors.Storage();
        var zip = FileCell.Read(storage).Root.ToArray();
        var update = FileCell.Write(zip, null, Ids(), ChunkingOptions.Original);
        var request = FssHttpBRequest.Read(FssHttpDVectors.PutChangesRequestBytes);
        await Assert.That(Canonical(update.Added)).IsEqualTo(Canonical(request.DataElements!.Elements));
        await Assert.That(update.Storage.Elements.Count).IsEqualTo(update.Added.Count);
    }

    public static IEnumerable<string> Files() =>
    [
        "Empty",
        "OneByte",
        "Text",
        "Simple",
        "Zip",
        "ZipLargeItem",
        "ZipSubChunks"
    ];

    static byte[] File(string name) =>
        name switch
        {
            "Empty" => [],
            "OneByte" => [42],
            "Text" => Zips.Text("Hello, World"),
            "Simple" => Zips.Random(2 * megabyte + 100, 1),
            "Zip" => Zips.Create(("a.txt", Zips.Text("a")), ("b.bin", Zips.Random(10_000, 2)), ("c.txt", Zips.Text("c"))),
            "ZipLargeItem" => Zips.Create(CompressionLevel.NoCompression, ("a.bin", Zips.Random(megabyte + 10, 3))),
            "ZipSubChunks" => Zips.Create(CompressionLevel.NoCompression, ("a.bin", Zips.Random(7 * megabyte, 4)), ("b.txt", Zips.Text("b"))),
            _ => throw new ArgumentException(name)
        };

    /// <summary>
    /// A file written into a new cell storage reads back the same, and its data elements go through the codec.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(Files))]
    public async Task Write_Read(string name)
    {
        var file = File(name);
        var update = FileCell.Write(file, null, Ids(), ChunkingOptions.Default);
        var tree = FileCell.Read(update.Storage);
        await AssertBytes(tree.Root.ToArray(), file);
        await Assert.That(update.Storage.Elements.Count).IsEqualTo(update.Added.Count);

        var put = new PutChangesRequest(1, update.Storage.StorageIndexId, ExtendedGuid.Null, 0);
        var request = new FssHttpBRequest(null, [put], new(update.Added));
        var read = FssHttpBRequest.Read(request.ToArray());
        await Assert.That(FssHttpB.Describe(read.DataElements)).IsEqualTo(FssHttpB.Describe(request.DataElements));
        var storage = CellStorage.Create(read.DataElements!.Elements, update.Storage.StorageIndexId);
        await AssertBytes(FileCell.Read(storage).Root.ToArray(), file);
    }

    /// <summary>
    /// 2.4.1: a chunk larger than 3 megabytes is an Intermediate Node Object with its signature, over a Leaf Node Object
    /// for each subchunk.
    /// </summary>
    [Test]
    public async Task Write_SubChunks()
    {
        var update = FileCell.Write(File("ZipSubChunks"), null, Ids(), ChunkingOptions.Default);
        var root = (IntermediateNode) FileCell.Read(update.Storage).Root;
        var data = (IntermediateNode) root.Children.Single(_ => _ is IntermediateNode);
        await Assert.That(data.Signature.Length).IsEqualTo(20);
        await Assert.That(data.Children.Select(_ => _.DataSize).ToList()).IsEquivalentTo([3ul * megabyte, 3ul * megabyte, (ulong) megabyte]);
        await Assert.That(data.Children.All(_ => _ is LeafNode { Signature.Length: 8, Children: [DataNode] })).IsTrue();
    }

    /// <summary>
    /// The same bytes again add nothing.
    /// </summary>
    [Test]
    public async Task Write_Unchanged()
    {
        var file = File("Zip");
        var first = FileCell.Write(file, null, Ids(), ChunkingOptions.Default);
        var second = FileCell.Write(file.ToArray(), first.Storage, Ids(), ChunkingOptions.Default);
        await Assert.That(second.Added.Count).IsEqualTo(0);
        await Assert.That(second.Storage).IsSameReferenceAs(first.Storage);
    }

    /// <summary>
    /// A zip with one item changed: the chunks of the other items are kept with their objects, so the new revision has
    /// new Object Groups for the new root, the changed item and the central directory only, and lists those of the kept
    /// chunks again, with no base revision. The Storage Index and Storage Manifest keep their identifiers, the Cell Manifest does not,
    /// and the data elements only the previous revision used go.
    /// </summary>
    [Test]
    public async Task Write_OneItemChanged()
    {
        var ids = Ids();
        var large = Zips.Random(10_000, 5);
        var before = Zips.Create(("a.txt", Zips.Text("a")), ("b.txt", Zips.Text("b")), ("c.bin", large));
        var after = Zips.Create(("a.txt", Zips.Text("a")), ("b.txt", Zips.Text("B")), ("c.bin", large));
        var first = FileCell.Write(before, null, ids, ChunkingOptions.Default);
        var second = FileCell.Write(after, first.Storage, ids, ChunkingOptions.Default);

        await AssertBytes(FileCell.Read(second.Storage).Root.ToArray(), after);
        await AssertBytes(FileCell.Read(first.Storage).Root.ToArray(), before);
        // the root, b.txt and the central directory, each a node and its data, but the root
        await Assert.That(second.Added.OfType<ObjectGroup>().Count()).IsEqualTo(5);
        await Assert.That(second.Added.Select(_ => _.Type).Distinct().ToList())
            .IsEquivalentTo([DataElementType.ObjectGroup, DataElementType.CellManifest, DataElementType.RevisionManifest, DataElementType.StorageIndex]);
        await Assert.That(second.Storage.StorageIndexId).IsEqualTo(first.Storage.StorageIndexId);

        var firstTree = FileCell.Read(first.Storage);
        var secondTree = FileCell.Read(second.Storage);
        var revision = second.Storage.GetRevision(secondTree.RevisionId);
        await Assert.That(revision.BaseRevisionId.IsNull).IsTrue();
        // the 5 new, and the leaf and data of a.txt, of the header of c.bin and of its data
        await Assert.That(revision.ObjectGroups.Count).IsEqualTo(11);
        await Assert.That(revision.ObjectGroups.Count(first.Storage.Elements.ContainsKey)).IsEqualTo(6);
        await Assert.That(second.Storage.Elements.Count).IsEqualTo(11 + 4);
        var firstCell = first.Added.OfType<CellManifest>().Single();
        var secondCell = second.Added.OfType<CellManifest>().Single();
        // a new Cell Manifest each revision, so that a client that expects the previous one fails coherency
        await Assert.That(secondCell.Id).IsNotEqualTo(firstCell.Id);
        await Assert.That(secondCell.SerialNumber).IsNotEqualTo(firstCell.SerialNumber);

        var kept = ((IntermediateNode) secondTree.Root).Children.Select(_ => _.Id)
            .Intersect(((IntermediateNode) firstTree.Root).Children.Select(_ => _.Id))
            .Count();
        // a.txt, and the header and data of c.bin
        await Assert.That(kept).IsEqualTo(3);
    }

    /// <summary>
    /// A chunk with subchunks is kept whole, with the objects of its subchunks.
    /// </summary>
    [Test]
    public async Task Write_SubChunksKept()
    {
        var ids = Ids();
        var large = Zips.Random(4 * megabyte, 8);
        var before = Zips.Create(CompressionLevel.NoCompression, ("a.txt", Zips.Text("a")), ("large.bin", large));
        var after = Zips.Create(CompressionLevel.NoCompression, ("a.txt", Zips.Text("A")), ("large.bin", large));
        var first = FileCell.Write(before, null, ids, ChunkingOptions.Default);
        var second = FileCell.Write(after, first.Storage, ids, ChunkingOptions.Default);
        await AssertBytes(FileCell.Read(second.Storage).Root.ToArray(), after);
        // the root, a.txt and the central directory, each a node and its data, but the root
        await Assert.That(second.Added.OfType<ObjectGroup>().Count()).IsEqualTo(5);
        var firstLarge = FileCell.Read(first.Storage).Root.Children.OfType<IntermediateNode>().Single();
        var secondLarge = FileCell.Read(second.Storage).Root.Children.OfType<IntermediateNode>().Single();
        await Assert.That(secondLarge.Id).IsEqualTo(firstLarge.Id);
        await Assert.That(secondLarge.Children.Select(_ => _.Id).ToList()).IsEquivalentTo(firstLarge.Children.Select(_ => _.Id).ToList());
    }

    /// <summary>
    /// A file with nothing in common with the previous one has no base revision, and the previous data elements go, but
    /// the Storage Manifest.
    /// </summary>
    [Test]
    public async Task Write_NothingKept()
    {
        var ids = Ids();
        var first = FileCell.Write(File("Zip"), null, ids, ChunkingOptions.Default);
        var second = FileCell.Write(File("Simple"), first.Storage, ids, ChunkingOptions.Default);
        var revision = second.Storage.GetRevision(FileCell.Read(second.Storage).RevisionId);
        await Assert.That(revision.BaseRevisionId.IsNull).IsTrue();
        await Assert.That(second.Storage.Elements.Count).IsEqualTo(second.Added.Count + 1);
        await Assert.That(second.Storage.Elements.Values.OfType<StorageManifest>().Single()).IsSameReferenceAs(first.Added.OfType<StorageManifest>().Single());
    }

    /// <summary>
    /// Revisions one after the other: only the current one and its data elements stay, and the chunks that do not change
    /// keep their objects throughout.
    /// </summary>
    [Test]
    public async Task Write_ManyRevisions()
    {
        var ids = Ids();
        var large = Zips.Random(10_000, 6);
        CellStorage? storage = null;
        var kept = new HashSet<string>();
        for (var version = 0; version < 12; version++)
        {
            var zip = Zips.Create(("version.txt", Zips.Text($"{version}")), ("large.bin", large));
            storage = FileCell.Write(zip, storage, ids, ChunkingOptions.Default).Storage;
            var root = FileCell.Read(storage).Root;
            await AssertBytes(root.ToArray(), zip);
            await Assert.That(storage.StorageIndex.Entries.OfType<StorageIndexRevisionMapping>().Count()).IsEqualTo(1);
            // the root and the leaf and data of each chunk, and the manifests and Storage Index
            await Assert.That(storage.Elements.Count).IsEqualTo(1 + 4 * 2 + 4);
            kept.Add($"{root.Children[1].Id} {root.Children[2].Id}");
        }

        // the header and data of large.bin
        await Assert.That(kept.Count).IsEqualTo(1);
    }

    /// <summary>
    /// The other roots of the previous revision, streams opaque to [MS-FSSHTTPD] like the one Word adds, are kept with
    /// their objects. One whose object is missing fails.
    /// </summary>
    [Test]
    public async Task Write_OtherRoots()
    {
        var ids = Ids();
        var first = FileCell.Write(File("Zip"), null, ids, ChunkingOptions.Default);
        var (leaf, data) = (ids.NewId(), ids.NewId());
        var otherRoot = new ExtendedGuid(FssHttpB.Guid("EDF48BAC-3745-4416-A4C9-D5B71BB870F6"), 2);
        var group = new ObjectGroup(
            ids.NewId(),
            ids.NewSerialNumber(),
            null,
            [new ObjectDeclaration(leaf, 1, 16, 1, 0), new ObjectDeclaration(data, 1, 3, 0, 0)],
            null,
            [new ObjectData([data], [], Node(true, 3)), new ObjectData([], [], Zips.Text("abc"))]);
        var revision = first.Storage.GetRevision(FileCell.Read(first.Storage).RevisionId);
        var withOther = Replace(first.Storage, group, revision with
        {
            Roots = [..revision.Roots, new(otherRoot, leaf)],
            ObjectGroups = [..revision.ObjectGroups, group.Id]
        });

        var second = FileCell.Write(File("Simple"), withOther, ids, ChunkingOptions.Default);
        var secondRevision = second.Storage.GetRevision(FileCell.Read(second.Storage).RevisionId);
        await Assert.That(secondRevision.Roots.Select(_ => _.RootId).ToList()).IsEquivalentTo([FileCell.RootId, otherRoot]);
        await Assert.That(secondRevision.ObjectGroups).Contains(group.Id);
        await Assert.That(second.Storage.Elements.ContainsKey(group.Id)).IsTrue();

        // a root on a chunk that the file keeps: its Object Groups are listed once
        var before = Zips.Create(("a.txt", Zips.Text("a")), ("b.txt", Zips.Text("b")));
        var after = Zips.Create(("a.txt", Zips.Text("a")), ("b.txt", Zips.Text("B")));
        var zip = FileCell.Write(before, null, ids, ChunkingOptions.Default);
        var zipTree = FileCell.Read(zip.Storage);
        var zipRevision = zip.Storage.GetRevision(zipTree.RevisionId);
        var shared = Replace(zip.Storage, zipRevision with
        {
            Roots = [..zipRevision.Roots, new(otherRoot, zipTree.Root.Children[0].Id)]
        });
        var sharedRevision = FileCell.Write(after, shared, ids, ChunkingOptions.Default).Added.OfType<RevisionManifest>().Single();
        await Assert.That(sharedRevision.ObjectGroups.Distinct().Count()).IsEqualTo(sharedRevision.ObjectGroups.Count);
        await Assert.That(sharedRevision.Roots[1].ObjectId).IsEqualTo(zipTree.Root.Children[0].Id);

        var broken = Replace(first.Storage, revision with
        {
            Roots = [..revision.Roots, new(otherRoot, ids.NewId())]
        });
        var exception = Assert.Throws<FssHttpDException>(() => FileCell.Write(File("Simple"), broken, ids, ChunkingOptions.Default));
        await Assert.That(exception.Message).Contains("The previous revision has no object");
    }

    /// <summary>
    /// Two chunks with the same bytes: the next revision keeps the first one's objects for one chunk only, as each object
    /// is in the tree once.
    /// </summary>
    [Test]
    public async Task Write_SameChunkTwice()
    {
        var ids = Ids();
        var half = Zips.Random(megabyte, 7);
        byte[] before = [..half, ..half];
        byte[] after = [..half, ..half, 1];
        var first = FileCell.Write(before, null, ids, ChunkingOptions.Default);
        var second = FileCell.Write(after, first.Storage, ids, ChunkingOptions.Default);
        var tree = FileCell.Read(second.Storage);
        await AssertBytes(tree.Root.ToArray(), after);
        // the root, then a new leaf and data for the second half and for the last byte
        await Assert.That(second.Added.OfType<ObjectGroup>().Count()).IsEqualTo(5);
    }

    static ObjectGroup Group(IdAllocator ids, ExtendedGuid objectId, ExtendedGuid[] references, ReadOnlyMemory<byte> data) =>
        new(
            ids.NewId(),
            ids.NewSerialNumber(),
            null,
            [new ObjectDeclaration(objectId, 1, (ulong) data.Length, (ulong) references.Length, 0)],
            null,
            [new ObjectData(references, [], data)]);

    static byte[] Node(bool leaf, ulong size, byte[]? signature = null) => new NodeObjectData(leaf, signature ?? [], size).ToArray();

    /// <summary>
    /// A cell storage of objects written by hand: a revision whose root is <paramref name="root"/>, with the Object Groups
    /// <paramref name="groups"/>, and other data elements that it does not list.
    /// </summary>
    static CellStorage Cell(IdAllocator ids, ExtendedGuid root, IEnumerable<ObjectGroup> groups, IEnumerable<DataElement>? others = null, RevisionManifest? baseManifest = null)
    {
        var listed = groups.ToList();
        var elements = new List<DataElement>([..listed, ..others ?? []]);
        var revisionId = ids.NewId();
        var revision = new RevisionManifest(ids.NewId(), ids.NewSerialNumber(), revisionId, baseManifest?.RevisionId ?? ExtendedGuid.Null, [new(FileCell.RootId, root)], listed.Select(_ => _.Id).ToList());
        var manifest = new StorageManifest(ids.NewId(), ids.NewSerialNumber(), FileCell.SchemaId, [new(FileCell.RootId, FileCell.CellId)]);
        var cell = new CellManifest(ids.NewId(), ids.NewSerialNumber(), revisionId);
        var entries = new List<StorageIndexEntry>
        {
            new StorageIndexManifestMapping(manifest.Id, ids.NewSerialNumber()),
            new StorageIndexCellMapping(FileCell.CellId, cell.Id, ids.NewSerialNumber()),
            new StorageIndexRevisionMapping(revisionId, revision.Id, ids.NewSerialNumber())
        };
        if (baseManifest != null)
        {
            entries.Add(new StorageIndexRevisionMapping(baseManifest.RevisionId, baseManifest.Id, ids.NewSerialNumber()));
            elements.Add(baseManifest);
        }

        var index = new StorageIndex(ids.NewId(), ids.NewSerialNumber(), entries);
        return CellStorage.Create([..elements, manifest, cell, revision, index], index.Id);
    }

    /// <summary>
    /// A Leaf Node Object as the root, as SharePoint Server 2010 returns a file of one chunk (the note &lt;1&gt;).
    /// </summary>
    [Test]
    public async Task Read_LeafRoot()
    {
        var ids = Ids();
        var leaf = ids.NewId();
        var data = ids.NewId();
        var storage = Cell(ids, leaf, [Group(ids, leaf, [data], Node(true, 3)), Group(ids, data, [], Zips.Text("abc"))]);
        var tree = FileCell.Read(storage);
        await Assert.That(tree.Root).IsTypeOf<LeafNode>();
        await AssertBytes(tree.Root.ToArray(), Zips.Text("abc"));
    }

    /// <summary>
    /// A Leaf Node Object over the Leaf Node Objects of subchunks (2.2.3.2), and objects in one Object Group.
    /// </summary>
    [Test]
    public async Task Read_LeafOverLeaves()
    {
        var ids = Ids();
        var (root, leaf, first, second, firstData, secondData) = (ids.NewId(), ids.NewId(), ids.NewId(), ids.NewId(), ids.NewId(), ids.NewId());
        var group = new ObjectGroup(
            ids.NewId(),
            ids.NewSerialNumber(),
            null,
            [
                new ObjectDeclaration(root, 1, 16, 1, 0),
                new ObjectDeclaration(leaf, 1, 16, 2, 0),
                new ObjectDeclaration(first, 1, 16, 1, 0),
                new ObjectDeclaration(second, 1, 16, 1, 0)
            ],
            null,
            [
                new ObjectData([leaf], [], Node(false, 6)),
                new ObjectData([first, second], [], Node(true, 6, [1, 2])),
                new ObjectData([firstData], [], Node(true, 3)),
                new ObjectData([secondData], [], Node(true, 3))
            ]);
        var storage = Cell(ids, root, [group, Group(ids, firstData, [], Zips.Text("abc")), Group(ids, secondData, [], Zips.Text("def"))]);
        var tree = FileCell.Read(storage);
        await AssertBytes(tree.Root.ToArray(), Zips.Text("abcdef"));
        var node = (LeafNode) ((IntermediateNode) tree.Root).Children.Single();
        await Assert.That(node.Children.All(_ => _ is LeafNode)).IsTrue();
    }

    /// <summary>
    /// The layout of Word's saves to SharePoint Server (docs/fsshttp-notes.md): each leaf with its Data Node Object in one
    /// Object Group, with Object Group metadata; leaves with a Data Hash of 20 bytes; data elements with null Serial
    /// Numbers; and a second root, an opaque stream that is a single leaf with no signature.
    /// </summary>
    [Test]
    public async Task Read_WordLayout()
    {
        var ids = Ids();
        var (root, leaf, data, streamLeaf, streamData) = (ids.NewId(), ids.NewId(), ids.NewId(), ids.NewId(), ids.NewId());
        var hash = new byte[20];
        ObjectGroup Pair(ExtendedGuid leafId, ExtendedGuid dataId, byte[] signature, string text) =>
            new(
                ids.NewId(),
                SerialNumber.Null,
                null,
                [new ObjectDeclaration(leafId, 1, 59, 1, 0), new ObjectDeclaration(dataId, 1, (ulong) text.Length, 0, 0)],
                [2, 2],
                [
                    new ObjectData([dataId], [], new NodeObjectData(true, signature, (ulong) text.Length, hash).ToArray()),
                    new ObjectData([], [], Zips.Text(text))
                ]);
        var rootGroup = new ObjectGroup(ids.NewId(), SerialNumber.Null, null, [new ObjectDeclaration(root, 1, 16, 1, 0)], [1], [new ObjectData([leaf], [], Node(false, 3))]);
        var storage = Cell(ids, root, [rootGroup, Pair(leaf, data, new byte[20], "abc"), Pair(streamLeaf, streamData, [], "stream")]);
        var revision = storage.Elements.Values.OfType<RevisionManifest>().Single();
        storage = Replace(storage, revision with
        {
            Roots = [new(new(FssHttpB.Guid("EDF48BAC-3745-4416-A4C9-D5B71BB870F6"), 2), streamLeaf), ..revision.Roots]
        });

        var tree = FileCell.Read(storage);
        await AssertBytes(tree.Root.ToArray(), Zips.Text("abc"));
        var node = (LeafNode) tree.Root.Children.Single();
        await Assert.That(node.DataHash!.Value.Length).IsEqualTo(20);

        // a write keeps the stream
        var update = FileCell.Write(Zips.Text("abcd"), storage, ids, ChunkingOptions.Default);
        await AssertBytes(FileCell.Read(update.Storage).Root.ToArray(), Zips.Text("abcd"));
        await Assert.That(update.Added.OfType<RevisionManifest>().Single().Roots.Count).IsEqualTo(2);
    }

    /// <summary>
    /// The data of a Data Node Object in an Object Data BLOB ([MS-FSSHTTPB] 2.2.1.12.6.2, 2.2.1.12.8), which pruning
    /// keeps.
    /// </summary>
    [Test]
    public async Task Read_Blob()
    {
        var ids = Ids();
        var (root, leaf, data, blob) = (ids.NewId(), ids.NewId(), ids.NewId(), ids.NewId());
        var blobGroup = new ObjectGroup(
            ids.NewId(),
            ids.NewSerialNumber(),
            null,
            [new ObjectDataBlobDeclaration(data, blob, 1, 0, 0)],
            null,
            [new ObjectDataBlobReference([], [], blob)]);
        var storage = Cell(ids, root, [Group(ids, root, [leaf], Node(false, 3)), Group(ids, leaf, [data], Node(true, 3)), blobGroup], [new ObjectDataBlob(blob, ids.NewSerialNumber(), Zips.Text("abc"))]);
        await AssertBytes(FileCell.Read(storage).Root.ToArray(), Zips.Text("abc"));
        var pruned = storage.Prune();
        await Assert.That(pruned.Elements.ContainsKey(blob)).IsTrue();
        await Assert.That(pruned.Elements.Count).IsEqualTo(storage.Elements.Count);
        await AssertBytes(FileCell.Read(pruned).Root.ToArray(), Zips.Text("abc"));
    }

    /// <summary>
    /// Objects of a base revision, which the revision does not declare again.
    /// </summary>
    [Test]
    public async Task Read_BaseRevision()
    {
        var ids = Ids();
        var (root, leaf, data) = (ids.NewId(), ids.NewId(), ids.NewId());
        var dataGroup = Group(ids, data, [], Zips.Text("abc"));
        var baseRevision = new RevisionManifest(ids.NewId(), ids.NewSerialNumber(), ids.NewId(), ExtendedGuid.Null, [], [dataGroup.Id]);
        var storage = Cell(ids, root, [Group(ids, root, [leaf], Node(false, 3)), Group(ids, leaf, [data], Node(true, 3))], [dataGroup], baseRevision);
        await AssertBytes(FileCell.Read(storage).Root.ToArray(), Zips.Text("abc"));
        await Assert.That(storage.Prune().Elements.ContainsKey(dataGroup.Id)).IsTrue();
    }

    static CellStorage Replace(CellStorage storage, params DataElement[] elements) => storage.With(elements, storage.StorageIndexId);

    static ObjectGroup GroupOf(CellStorage storage, ExtendedGuid objectId) =>
        storage.Elements.Values.OfType<ObjectGroup>().Single(_ => _.Declarations.Any(_ => _.ObjectId == objectId));

    /// <summary>
    /// A cell storage written from a zip of two items, its tree, and its revision.
    /// </summary>
    static (CellStorage Storage, IntermediateNode Root, RevisionManifest Revision) Sample()
    {
        var update = FileCell.Write(Zips.Create(("a.txt", Zips.Text("a")), ("b.txt", Zips.Text("b"))), null, Ids(), ChunkingOptions.Default);
        var tree = FileCell.Read(update.Storage);
        return (update.Storage, (IntermediateNode) tree.Root, update.Storage.GetRevision(tree.RevisionId));
    }

    public static IEnumerable<string> InvalidCases() => Invalid().Select(_ => _.Name);

    static IEnumerable<(string Name, Func<CellStorage> Storage, string Message)> Invalid()
    {
        yield return ("NoManifestMapping", () =>
        {
            var (storage, _, _) = Sample();
            return Replace(storage, storage.StorageIndex with
            {
                Entries = storage.StorageIndex.Entries.Where(_ => _ is not StorageIndexManifestMapping).ToList()
            });
        }, "no Storage Manifest mapping");
        yield return ("Schema", () =>
        {
            var (storage, _, _) = Sample();
            var manifest = storage.Elements.Values.OfType<StorageManifest>().Single();
            return Replace(storage, manifest with
            {
                SchemaId = Guid.NewGuid()
            });
        }, "not the one of a file");
        yield return ("NoStorageRoot", () =>
        {
            var (storage, _, _) = Sample();
            var manifest = storage.Elements.Values.OfType<StorageManifest>().Single();
            return Replace(storage, manifest with
            {
                Roots = []
            });
        }, "The Storage Manifest has no root");
        yield return ("NoCellMapping", () =>
        {
            var (storage, _, _) = Sample();
            return Replace(storage, storage.StorageIndex with
            {
                Entries = storage.StorageIndex.Entries.Where(_ => _ is not StorageIndexCellMapping).ToList()
            });
        }, "no mapping for the cell");
        yield return ("NoRevisionMapping", () =>
        {
            var (storage, _, _) = Sample();
            return Replace(storage, storage.StorageIndex with
            {
                Entries = storage.StorageIndex.Entries.Where(_ => _ is not StorageIndexRevisionMapping).ToList()
            });
        }, "No revision mapping");
        yield return ("NoRevisionRoot", () =>
        {
            var (storage, _, revision) = Sample();
            return Replace(storage, revision with
            {
                Roots = []
            });
        }, "has no root");
        yield return ("MissingElement", () =>
        {
            var (storage, _, revision) = Sample();
            var elements = storage.Elements.Where(_ => _.Key != revision.ObjectGroups[1]).ToDictionary();
            return storage with
            {
                Elements = elements
            };
        }, "No data element");
        yield return ("WrongType", () =>
        {
            var (storage, _, revision) = Sample();
            return Replace(storage, revision with
            {
                ObjectGroups = [..revision.ObjectGroups, storage.StorageIndexId]
            });
        }, "is a StorageIndex, not a ObjectGroup");
        yield return ("DeclarationsAndData", () =>
        {
            var (storage, _, revision) = Sample();
            var group = storage.Get<ObjectGroup>(revision.ObjectGroups[0]);
            return Replace(storage, group with
            {
                Data = []
            });
        }, "1 declarations and 0 data");
        yield return ("MissingObject", () =>
        {
            var (storage, _, revision) = Sample();
            return Replace(storage, revision with
            {
                ObjectGroups = revision.ObjectGroups.Skip(1).ToList()
            });
        }, "The revision has no object");
        yield return ("ObjectTwice", () =>
        {
            var (storage, root, _) = Sample();
            var group = GroupOf(storage, root.Id);
            var data = (ObjectData) group.Data[0];
            return Replace(storage, group with
            {
                Data = [data with { ObjectReferences = [..data.ObjectReferences, data.ObjectReferences[0]] }]
            });
        }, "in the tree more than once");
        yield return ("SizeMismatch", () =>
        {
            var (storage, root, _) = Sample();
            var group = GroupOf(storage, root.Id);
            var data = (ObjectData) group.Data[0];
            return Replace(storage, group with
            {
                Data = [data with { Data = Node(false, root.DataSize + 1) }]
            });
        }, "its children");
        yield return ("ChildrenTooLarge", () =>
        {
            var (storage, root, _) = Sample();
            var group = GroupOf(storage, root.Id);
            var data = (ObjectData) group.Data[0];
            return Replace(storage, group with
            {
                Data = [data with { Data = Node(false, 1) }]
            });
        }, "have more than its 1 bytes");
        yield return ("NotANode", () =>
        {
            var (storage, root, _) = Sample();
            var group = GroupOf(storage, root.Id);
            var data = (ObjectData) group.Data[0];
            return Replace(storage, group with
            {
                Data = [data with { Data = Zips.Text("PK") }]
            });
        }, "is not a node object");
        yield return ("LeafMixed", () =>
        {
            var (storage, root, _) = Sample();
            var leaves = root.Children.Cast<LeafNode>().ToList();
            var group = GroupOf(storage, leaves[0].Id);
            var data = (ObjectData) group.Data[0];
            return Replace(storage, group with
            {
                Data = [data with { ObjectReferences = [..data.ObjectReferences, leaves[1].Id], Data = Node(true, leaves[0].DataSize + leaves[1].DataSize) }]
            });
        }, "neither a single Data Node Object nor Leaf Node Objects");
        yield return ("Excluded", () =>
        {
            var (storage, root, _) = Sample();
            var dataNode = ((LeafNode) root.Children[0]).Children[0];
            var group = GroupOf(storage, dataNode.Id);
            return Replace(storage, group with
            {
                Data = [new ObjectExcludedData([], [], dataNode.DataSize)]
            });
        }, "is excluded");
        yield return ("RevisionLoop", () =>
        {
            var (storage, _, revision) = Sample();
            return Replace(storage, revision with
            {
                BaseRevisionId = revision.RevisionId
            });
        }, "loop");
        yield return ("TooDeep", () =>
        {
            var ids = Ids();
            var nodes = Enumerable.Range(0, 20).Select(_ => ids.NewId()).ToList();
            var data = ids.NewId();
            var groups = new List<DataElement>();
            for (var index = 0; index < nodes.Count; index++)
            {
                ExtendedGuid[] references = index + 1 < nodes.Count ? [nodes[index + 1]] : [data];
                groups.Add(Group(ids, nodes[index], references, Node(index + 1 == nodes.Count, 1)));
            }

            groups.Add(Group(ids, data, [], new byte[1]));
            return Cell(ids, nodes[0], groups.Cast<ObjectGroup>());
        }, "deeper than 16");
    }

    /// <summary>
    /// Data elements that are not a file stream, each failing with what is wrong.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(InvalidCases))]
    public async Task Read_Invalid(string name)
    {
        var (_, storage, message) = Invalid().Single(_ => _.Name == name);
        var exception = Assert.Throws<FssHttpDException>(() => FileCell.Read(storage()));
        await Assert.That(exception.Message).Contains(message);
    }

    [Test]
    public async Task ToArray_TooLarge()
    {
        var node = new IntermediateNode(ExtendedGuid.Null, ReadOnlyMemory<byte>.Empty, 3UL * 1024 * 1024 * 1024, []);
        await Assert.That(node.ToArray).Throws<FssHttpDException>();
    }

    /// <summary>
    /// [MS-FSSHTTPD] 2.2.2.1 and 2.2.3.1: the node objects of the example, read and written back; a leaf with a Data Hash;
    /// a signature too long for a 16-bit header.
    /// </summary>
    [Test]
    [Arguments("0401" + "080300" + "1011DC00000000000000" + "81")]
    [Arguments("FC00" + "0853" + "51F333D2A6BB6F43C9817AAB3A629D3C8A395F109D8289D1F705000000000000000500000000000000" + "10112C00000000000000" + "7D")]
    [Arguments("FC00" + "080503AB" + "10110300000000000000" + "780705CDEF" + "7D")]
    public async Task NodeObjectData_RoundTrip(string hex)
    {
        var bytes = FssHttpB.Hex(hex);
        var node = NodeObjectData.Read(bytes);
        await Assert.That(FssHttpB.Hex(node.ToArray())).IsEqualTo(hex);
    }

    [Test]
    public async Task NodeObjectData_Values()
    {
        var hash = new NodeObjectData(true, new byte[200], 3, new byte[] { 1, 2 });
        var read = NodeObjectData.Read(hash.ToArray());
        await Assert.That(read.Leaf).IsTrue();
        await Assert.That(read.Signature.Length).IsEqualTo(200);
        await Assert.That(read.DataSize).IsEqualTo(3ul);
        await Assert.That(FssHttpB.Hex(read.DataHash!.Value.Span)).IsEqualTo("0102");
        var intermediate = NodeObjectData.Read(FssHttpB.Hex("0401" + "080300" + "10110600000000000000" + "81"));
        await Assert.That(intermediate.Leaf).IsFalse();
        await Assert.That(intermediate.DataHash).IsNull();
    }

    /// <summary>
    /// Not a node object: another start, a Data Hash in an intermediate node, bytes after the end.
    /// </summary>
    [Test]
    [Arguments("8400")]
    [Arguments("0401" + "080300" + "10110000000000000000" + "78050300" + "81")]
    [Arguments("0401" + "080300" + "10110000000000000000" + "81" + "00")]
    [Arguments("0401" + "080300" + "81")]
    public async Task NodeObjectData_Invalid(string hex) =>
        await Assert.That(() => NodeObjectData.Read(FssHttpB.Hex(hex))).Throws<FssHttpBException>();

    [Test]
    public async Task CellStorage_Get()
    {
        var (storage, _, revision) = Sample();
        await Assert.That(storage.Get<RevisionManifest>(revision.Id)).IsSameReferenceAs(revision);
        await Assert.That(() => storage.Get<CellManifest>(revision.Id)).Throws<FssHttpDException>();
        await Assert.That(() => storage.Get<CellManifest>(ExtendedGuid.Null)).Throws<FssHttpDException>();
        // a later version of a data element replaces it
        var later = revision with
        {
            SerialNumber = new(Guid.NewGuid(), 1)
        };
        await Assert.That(Replace(storage, later).Get<RevisionManifest>(revision.Id)).IsSameReferenceAs(later);
    }

    /// <summary>
    /// Pruning keeps what it reaches, and passes over a listed Object Group that is missing.
    /// </summary>
    [Test]
    public async Task Prune_MissingGroup()
    {
        var (storage, _, revision) = Sample();
        var missing = revision.ObjectGroups[0];
        var elements = storage.Elements.Where(_ => _.Key != missing).ToDictionary();
        var pruned = (storage with
        {
            Elements = elements
        }).Prune();
        await Assert.That(pruned.Elements.Count).IsEqualTo(elements.Count);
    }

    [Test]
    public async Task IdAllocator_Values()
    {
        var guid = Guid.NewGuid();
        var ids = new IdAllocator(guid, uint.MaxValue);
        await Assert.That(ids.NewId()).IsEqualTo(new(guid, uint.MaxValue));
        await Assert.That(ids.NewSerialNumber()).IsEqualTo(new(guid, (ulong) uint.MaxValue + 1));
        await Assert.That(ids.Next).IsEqualTo((ulong) uint.MaxValue + 2);
        await Assert.That(ids.Guid).IsEqualTo(guid);
        await Assert.That(ids.NewId).Throws<InvalidOperationException>();
    }
}
