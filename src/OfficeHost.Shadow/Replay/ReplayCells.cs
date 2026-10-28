/// <summary>
/// The cells of the documents of a replay (docs/replay.md), to translate the Cell sub-requests that carry the ids of the
/// recorded server. For each document and partition:
/// - the copy the recorded server's clients have: what it sent them, and what they put;
/// - the copy a stand-in client has of OfficeHost's: what OfficeHost sent the replay, and what the replay put;
/// - the knowledge of each server paired, from the answers both gave to the same Cell: a query's knowledge of the
///   recorded server is replaced by OfficeHost's of the pair.
/// A put of the content is made again: the file the client put, read from its copy, is put against OfficeHost's copy,
/// as WordClient.Save puts. That needs the client's whole copy: a client that opened the document from Office's cache
/// before the recording put changes to data elements the recording does not have, and its put is sent as it was.
/// </summary>
class ReplayCells
{
    /// <summary>
    /// What a Cell sub-request became: its document and partition, whether it was translated, and the put sent to
    /// each server, to apply to the copies once they accepted it.
    /// </summary>
    public record Translation(
        string Document,
        Guid Partition,
        string? Status,
        IReadOnlyList<DataElement> RecordedPackage,
        ExtendedGuid? RecordedStorageIndex,
        IReadOnlyList<DataElement> ReplayPackage,
        ExtendedGuid? ReplayStorageIndex);

    public const string Translated = "translated";

    public const string FromCache = "Office's cache";

    sealed class Copies
    {
        public CellStorage? Recorded;
        public CellStorage? Replay;
        public Dictionary<string, Knowledge> Knowledge = [];
        public Knowledge? LatestReplayKnowledge;
    }

    Dictionary<(string Document, Guid Partition), Copies> cells = [];
    IdAllocator ids = new(Guid.NewGuid());

    Copies Get(string document, Guid partition)
    {
        if (!cells.TryGetValue((document, partition), out var copies))
        {
            copies = new();
            cells[(document, partition)] = copies;
        }

        return copies;
    }

    static string Key(Knowledge knowledge)
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        knowledge.Write(ref writer);
        return Convert.ToHexString(output.WrittenSpan);
    }

    /// <summary>
    /// The request with the knowledge of its queries, and a put of the content, translated to OfficeHost's, or null
    /// when nothing in it is the recorded server's. Assumes one partition per Cell, and a put at most, as Office sends.
    /// </summary>
    public byte[]? Translate(string document, Guid defaultPartition, FssHttpBRequest request, out Translation translation)
    {
        var first = request.SubRequests.Count > 0 ? request.SubRequests[0] : null;
        var partition = first?.TargetPartition ?? defaultPartition;
        var copies = Get(document, partition);
        var package = request.DataElements?.Elements ?? [];
        string? status = null;
        var changed = false;
        var subRequests = new List<SubRequest>();
        var replayPackage = package;
        ExtendedGuid? recordedIndex = null;
        ExtendedGuid? replayIndex = null;
        foreach (var subRequest in request.SubRequests)
        {
            switch (subRequest)
            {
                case QueryChangesRequest { Knowledge.Items.Count: > 0 } query:
                    var (knowledge, known) = Map(copies, query.Knowledge!);
                    status = Worst(status, known);
                    subRequests.Add(query with
                    {
                        Knowledge = knowledge
                    });
                    changed = true;
                    break;
                case PutChangesRequest put:
                    recordedIndex = put.StorageIndex;
                    replayIndex = put.StorageIndex;
                    var clientKnowledge = put.ClientKnowledge;
                    if (clientKnowledge is { Items.Count: > 0 })
                    {
                        clientKnowledge = Map(copies, clientKnowledge).Knowledge;
                    }

                    // a put of another partition, or of a whole file expecting nothing (a Save As): only the client's
                    // own ids, nothing to make again
                    if (partition != FssHttpCells.ContentPartition ||
                        (put.ExpectedStorageIndex.IsNull && copies.Recorded == null))
                    {
                        subRequests.Add(put with
                        {
                            ClientKnowledge = clientKnowledge
                        });
                        changed = true;
                        break;
                    }

                    if (Again(copies, put, package) is { } again)
                    {
                        replayPackage = again.Package;
                        replayIndex = again.Put.StorageIndex;
                        subRequests.Add(again.Put with
                        {
                            ClientKnowledge = clientKnowledge
                        });
                        status = Worst(status, Translated);
                    }
                    else
                    {
                        subRequests.Add(put);
                        status = FromCache;
                    }

                    changed = true;
                    break;
                default:
                    subRequests.Add(subRequest);
                    break;
            }
        }

        translation = new(document, partition, status, package, recordedIndex, replayPackage, replayIndex);
        if (!changed)
        {
            return null;
        }

        return (request with
        {
            SubRequests = subRequests,
            DataElements = new(replayPackage)
        }).ToArray();
    }

    static string Worst(string? status, string next)
    {
        if (status == FromCache)
        {
            return status;
        }

        return next;
    }

    /// <summary>
    /// OfficeHost's knowledge for the recorded server's: the one paired with it, else the stand-in's latest, else none
    /// (the recorded client knew the document from before the recording).
    /// </summary>
    static (Knowledge Knowledge, string Status) Map(Copies copies, Knowledge recorded)
    {
        if (copies.Knowledge.TryGetValue(Key(recorded), out var paired))
        {
            return (paired, Translated);
        }

        if (copies.LatestReplayKnowledge is { } latest)
        {
            return (latest, Translated);
        }

        return (Knowledge.Empty, FromCache);
    }

    /// <summary>
    /// The put made again against OfficeHost's copy, of the file the client put, or null when the client's copy does not
    /// have it all.
    /// </summary>
    (PutChangesRequest Put, IReadOnlyList<DataElement> Package)? Again(Copies copies, PutChangesRequest put, IReadOnlyList<DataElement> package)
    {
        if (copies.Recorded == null ||
            copies.Replay == null)
        {
            return null;
        }

        byte[] file;
        try
        {
            file = FileCell.Read(copies.Recorded.With(package, put.StorageIndex)).Root.ToArray();
        }
        catch (Exception exception) when (exception is FssHttpDException or FssHttpBException or InvalidDataException)
        {
            return null;
        }

        var replay = copies.Replay;
        var update = FileCell.Write(file, replay, ids, ChunkingOptions.Default);
        var index = update.Storage.StorageIndex with
        {
            Id = ids.NewId()
        };
        var expected = replay.StorageIndex with
        {
            Id = ids.NewId(),
            Entries = replay.StorageIndex.Entries.Select(Remap).ToList()
        };
        var again = update.Added
            .Where(_ => _ is not StorageIndex)
            .Append(index)
            .Append(expected)
            .Select(_ => _ with { SerialNumber = SerialNumber.Null })
            .ToList();
        return (put with
        {
            StorageIndex = index.Id,
            ExpectedStorageIndex = expected.Id
        }, again);
    }

    StorageIndexEntry Remap(StorageIndexEntry entry) =>
        entry switch
        {
            StorageIndexManifestMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            StorageIndexCellMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            StorageIndexRevisionMapping value => value with { MappingSerialNumber = ids.NewSerialNumber() },
            _ => entry
        };

    /// <summary>
    /// Applies the answers of each server to its copy, and pairs their knowledge.
    /// </summary>
    public void Answered(Translation translation, FssHttpBResponse? recorded, FssHttpBResponse? replay)
    {
        var copies = Get(translation.Document, translation.Partition);
        var recordedKnowledge = Apply(ref copies.Recorded, recorded, translation.RecordedPackage, translation.RecordedStorageIndex);
        var replayKnowledge = Apply(ref copies.Replay, replay, translation.ReplayPackage, translation.ReplayStorageIndex);
        if (replayKnowledge == null)
        {
            return;
        }

        copies.LatestReplayKnowledge = replayKnowledge;
        if (recordedKnowledge != null)
        {
            copies.Knowledge[Key(recordedKnowledge)] = replayKnowledge;
        }
    }

    static Knowledge? Apply(ref CellStorage? copy, FssHttpBResponse? response, IReadOnlyList<DataElement> package, ExtendedGuid? storageIndex)
    {
        Knowledge? knowledge = null;
        foreach (var subResponse in response?.SubResponses ?? [])
        {
            switch (subResponse)
            {
                case QueryChangesResponse query:
                    var elements = response!.DataElements?.Elements ?? [];
                    copy = copy?.With(elements, query.StorageIndex) ?? CellStorage.Create(elements, query.StorageIndex);
                    knowledge = query.Knowledge;
                    break;
                case PutChangesResponse put when storageIndex is { } index:
                    copy = copy?.With(package, index) ?? CellStorage.Create(package, index);
                    knowledge = put.Knowledge;
                    break;
            }
        }

        return knowledge;
    }

    /// <summary>
    /// The file of a copy of a partition (the document, the editors table, Office's metadata), or null when the copy
    /// does not have it all.
    /// </summary>
    public byte[]? File(string document, Guid partition, bool recorded)
    {
        if (!cells.TryGetValue((document, partition), out var copies))
        {
            return null;
        }

        var copy = recorded ? copies.Recorded : copies.Replay;
        if (copy == null)
        {
            return null;
        }

        try
        {
            return FileCell.Read(copy).Root.ToArray();
        }
        catch (Exception exception) when (exception is FssHttpDException or FssHttpBException or InvalidDataException or KeyNotFoundException)
        {
            return null;
        }
    }
}
