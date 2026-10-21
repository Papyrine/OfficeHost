/// <summary>
/// Serial Numbers as Cell Knowledge holds them ([MS-FSSHTTPB] 2.2.1.13.2): for each GUID, ranges of values.
/// </summary>
class SerialNumberSet
{
    Dictionary<Guid, List<(ulong From, ulong To)>> ranges = [];

    /// <summary>
    /// The Serial Numbers of the Cell Knowledge in a knowledge. Other kinds of knowledge are left out: this host keeps
    /// no waterlines or content tags.
    /// </summary>
    public static SerialNumberSet Of(Knowledge? knowledge)
    {
        var set = new SerialNumberSet();
        foreach (var item in knowledge?.Items.OfType<CellKnowledge>().SelectMany(_ => _.Items) ?? [])
        {
            switch (item)
            {
                case CellKnowledgeRange range:
                    set.Add(range.Guid, range.From, range.To);
                    break;
                case CellKnowledgeEntry entry:
                    set.Add(entry.SerialNumber);
                    break;
            }
        }

        return set;
    }

    public void Add(SerialNumber value)
    {
        if (!value.IsNull)
        {
            Add(value.Guid, value.Value, value.Value);
        }
    }

    void Add(Guid guid, ulong from, ulong to)
    {
        if (from > to)
        {
            return;
        }

        if (!ranges.TryGetValue(guid, out var list))
        {
            list = [];
            ranges[guid] = list;
        }

        list.Add((from, to));
    }

    public bool Contains(SerialNumber value) =>
        !value.IsNull &&
        ranges.TryGetValue(value.Guid, out var list) &&
        list.Any(_ => _.From <= value.Value && value.Value <= _.To);

    /// <summary>
    /// The set as Cell Knowledge: the ranges of each GUID merged.
    /// </summary>
    public Knowledge ToKnowledge()
    {
        var items = new List<CellKnowledgeItem>();
        foreach (var (guid, list) in ranges.OrderBy(_ => _.Key))
        {
            ulong? from = null;
            ulong to = 0;
            foreach (var range in list.OrderBy(_ => _.From))
            {
                if (from != null &&
                    (to == ulong.MaxValue || range.From <= to + 1))
                {
                    to = Math.Max(to, range.To);
                    continue;
                }

                if (from != null)
                {
                    items.Add(new CellKnowledgeRange(guid, from.Value, to));
                }

                from = range.From;
                to = range.To;
            }

            items.Add(new CellKnowledgeRange(guid, from!.Value, to));
        }

        return new([new CellKnowledge(items)]);
    }

    /// <summary>
    /// The knowledge of a whole cell storage: for each GUID of the Serial Numbers of its data elements and of the
    /// mappings of its Storage Index (2.2.1.12.2), the values from zero to the highest.
    /// </summary>
    public static Knowledge Full(CellStorage storage)
    {
        var highest = new Dictionary<Guid, ulong>();

        void Add(SerialNumber value)
        {
            if (!value.IsNull)
            {
                highest[value.Guid] = Math.Max(highest.GetValueOrDefault(value.Guid), value.Value);
            }
        }

        foreach (var element in storage.Elements.Values)
        {
            Add(element.SerialNumber);
        }

        foreach (var entry in storage.StorageIndex.Entries)
        {
            Add(entry.MappingSerialNumber);
        }

        var set = new SerialNumberSet();
        foreach (var (guid, value) in highest)
        {
            set.Add(guid, 0, value);
        }

        return set.ToKnowledge();
    }
}

/// <summary>
/// The result of a Put Changes: the sub-response, and the state to store when it succeeded.
/// </summary>
record PutOutcome(SubResponse Response, CellState? State, IReadOnlyList<DataElement> Added);

/// <summary>
/// Query Changes and Put Changes ([MS-FSSHTTPB] 3.1.4.2, 3.1.4.3) against the state of a cell storage.
/// </summary>
static class CellSync
{
    /// <summary>
    /// A cell storage with nothing in it: a Storage Index with a null Storage Manifest mapping, as SharePoint answers for
    /// a partition no client wrote.
    /// </summary>
    public static CellStorage Empty(IdAllocator ids)
    {
        var index = new StorageIndex(ids.NewId(), ids.NewSerialNumber(), [new StorageIndexManifestMapping(ExtendedGuid.Null, SerialNumber.Null)]);
        return CellStorage.Create([index], index.Id);
    }

    public static bool IsEmpty(CellStorage storage) =>
        storage.StorageIndex.Entries.All(_ => _.MappingId.IsNull);

    /// <summary>
    /// The data elements the client does not know of, by their Serial Numbers, up to the Maximum Data Elements of the
    /// <paramref name="query"/> in bytes (at least one), leaves first: Object Data BLOBs and Object Groups, then the manifests, then the Storage Index.
    /// When they do not all fit, the response is partial, and its knowledge is that of the client with the elements
    /// sent; otherwise it is the whole knowledge of the storage.
    /// </summary>
    public static (QueryChangesResponse Response, List<DataElement> Elements) Query(QueryChangesRequest query, CellStorage storage)
    {
        var known = SerialNumberSet.Of(query.Knowledge);
        var missing = storage.Elements.Values
            .Where(_ => !known.Contains(_.SerialNumber))
            .OrderBy(Order)
            .ToList();
        var limit = query.MaxDataElements ?? ulong.MaxValue;
        var sent = new List<DataElement>();
        ulong size = 0;
        foreach (var element in missing)
        {
            var elementSize = (ulong) Size(element);
            if (sent.Count > 0 &&
                elementSize > limit - Math.Min(limit, size))
            {
                break;
            }

            size += elementSize;
            sent.Add(element);
        }

        var partial = sent.Count < missing.Count;
        Knowledge knowledge;
        byte flags = 0;
        if (partial)
        {
            foreach (var element in sent)
            {
                known.Add(element.SerialNumber);
            }

            knowledge = known.ToKnowledge();
            flags = 1;
        }
        else
        {
            knowledge = SerialNumberSet.Full(storage);
        }

        return (new(query.RequestId, storage.StorageIndexId, flags, knowledge), sent);
    }

    static int Order(DataElement element) =>
        element.Type switch
        {
            DataElementType.ObjectDataBlob => 0,
            DataElementType.DataElementFragment => 0,
            DataElementType.ObjectGroup => 1,
            DataElementType.RevisionManifest => 2,
            DataElementType.CellManifest => 3,
            DataElementType.StorageManifest => 4,
            _ => 5
        };

    static int Size(DataElement element) => FssHttpB(element, element.SerialNumber).Length;

    static byte[] FssHttpB(DataElement element, SerialNumber serialNumber)
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        (element with
        {
            SerialNumber = serialNumber
        }).Write(ref writer);
        return output.WrittenSpan.ToArray();
    }

    /// <summary>
    /// Applies a Put Changes, as SharePoint Server SE does (docs/fsshttp-notes.md):
    /// - a Storage Index, or an expected Storage Index other than the current one, that is not in the package, or a data
    ///   element the new Storage Index reaches that neither the package nor the storage has, is a referenced data
    ///   element not found;
    /// - the mappings the new Storage Index updates have to be those the client expects, or none when it expects none
    ///   and implies null ([MS-FSSHTTPB] 2.2.2.1.4): otherwise it is a coherency failure, and the client queries the
    ///   latest changes and puts again;
    /// - with check for id reuse, a data element with the id of one the storage has, and other content, is an Extended
    ///   GUID collision.
    /// The data elements get Serial Numbers of the storage; one the storage has already, with the same content, is kept
    /// as it is. The storage keeps only what the new Storage Index reaches.
    /// </summary>
    public static PutOutcome Put(PutChangesRequest put, IReadOnlyList<DataElement> package, CellState state)
    {
        var current = state.Storage;
        var index = package.OfType<StorageIndex>().FirstOrDefault(_ => _.Id == put.StorageIndex);
        if (index == null)
        {
            return Failed(put, CellErrorCode.ReferencedDataElementNotFound);
        }

        StorageIndex? expected = null;
        if (!put.ExpectedStorageIndex.IsNull &&
            put.ExpectedStorageIndex != current?.StorageIndexId)
        {
            expected = package.OfType<StorageIndex>().FirstOrDefault(_ => _.Id == put.ExpectedStorageIndex);
            if (expected == null)
            {
                return Failed(put, CellErrorCode.ReferencedDataElementNotFound);
            }
        }

        if (current != null &&
            put.ExpectedStorageIndex != current.StorageIndexId &&
            !Coherent(put, index, expected, current.StorageIndex))
        {
            return Failed(put, CellErrorCode.CoherencyFailure);
        }

        var ids = state.Ids();
        var added = new List<DataElement>();
        foreach (var element in package)
        {
            if (current != null &&
                current.Elements.TryGetValue(element.Id, out var existing))
            {
                if (FssHttpB(existing, SerialNumber.Null).AsSpan().SequenceEqual(FssHttpB(element, SerialNumber.Null)))
                {
                    continue;
                }

                if (put.CheckForIdReuse)
                {
                    return Failed(put, CellErrorCode.ExtendedGuidCollision);
                }
            }

            added.Add(element with
            {
                SerialNumber = ids.NewSerialNumber()
            });
        }

        CellStorage storage;
        try
        {
            storage = (current ?? CellStorage.Create([], put.StorageIndex)).With(added, put.StorageIndex);
            if (storage.HasMissing())
            {
                return Failed(put, CellErrorCode.ReferencedDataElementNotFound);
            }

            storage = storage.Prune();
        }
        catch (FssHttpDException)
        {
            return Failed(put, CellErrorCode.ReferencedDataElementNotFound);
        }

        var response = new PutChangesResponse(
            put.RequestId,
            SerialNumberSet.Full(storage),
            true,
            ExtendedGuid.Null,
            added.Select(_ => _.Id).ToList());
        return new(response, new(storage, ids.Guid, ids.Next, state.FileVersion), added);
    }

    /// <summary>
    /// Whether the mappings the new Storage Index updates are as the client expects ([MS-FSSHTTPB] 2.2.2.1.4): for each
    /// of them, the storage maps it as the expected Storage Index does, or, when that has no such mapping and the client
    /// implies null, not at all. Word sends the Storage Index it expects in the package, under an identifier of its own,
    /// with the mappings of the Storage Index it got but Serial Numbers of its own (docs/fsshttp-notes.md): the targets
    /// are compared, not the Serial Numbers.
    /// </summary>
    static bool Coherent(PutChangesRequest put, StorageIndex index, StorageIndex? expected, StorageIndex current)
    {
        var server = Mappings(current);
        var expectedMappings = Mappings(expected);
        foreach (var key in Mappings(index).Keys)
        {
            var value = server.GetValueOrDefault(key, ExtendedGuid.Null);
            if (expectedMappings.TryGetValue(key, out var expectedValue))
            {
                if (expectedValue != value)
                {
                    return false;
                }

                continue;
            }

            if (put.ImplyNullExpectedIfNoMapping &&
                !value.IsNull)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// The targets of a Storage Index, by what they map: the Storage Manifest, a cell, or a revision.
    /// </summary>
    static Dictionary<object, ExtendedGuid> Mappings(StorageIndex? index)
    {
        var mappings = new Dictionary<object, ExtendedGuid>();
        foreach (var entry in index?.Entries ?? [])
        {
            object key = entry switch
            {
                StorageIndexCellMapping cell => cell.Cell,
                StorageIndexRevisionMapping revision => revision.Revision,
                _ => "manifest"
            };
            mappings[key] = entry.MappingId;
        }

        return mappings;
    }

    static PutOutcome Failed(PutChangesRequest put, CellErrorCode code) =>
        new(new FailedSubResponse(put.RequestId, RequestType.PutChanges, ResponseError.Cell(code)), null, []);
}
