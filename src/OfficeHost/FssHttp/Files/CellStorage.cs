/// <summary>
/// The data elements of a cell storage ([MS-FSSHTTPB] 3.1.1): the current version of each data element by its
/// identifier, and the identifier of the Storage Index (2.2.1.12.2), which they are reached from.
/// </summary>
record CellStorage(IReadOnlyDictionary<ExtendedGuid, DataElement> Elements, ExtendedGuid StorageIndexId)
{
    public static CellStorage Create(IEnumerable<DataElement> elements, ExtendedGuid storageIndexId) =>
        new(new Dictionary<ExtendedGuid, DataElement>().With(elements), storageIndexId);

    /// <summary>
    /// The storage with data elements added: a data element replaces the one with its identifier, as a later version
    /// of it, such as the Storage Index of the next revision.
    /// </summary>
    public CellStorage With(IEnumerable<DataElement> added, ExtendedGuid storageIndexId) =>
        new(new Dictionary<ExtendedGuid, DataElement>(Elements).With(added), storageIndexId);

    public StorageIndex StorageIndex => Get<StorageIndex>(StorageIndexId);

    public T Get<T>(ExtendedGuid id)
        where T : DataElement
    {
        if (!Elements.TryGetValue(id, out var element))
        {
            throw new FssHttpDException($"No data element {id}, for a {typeof(T).Name}");
        }

        if (element is not T result)
        {
            throw new FssHttpDException($"The data element {id} is a {element.Type}, not a {typeof(T).Name}");
        }

        return result;
    }

    /// <summary>
    /// The revision manifest of a revision, through the revision mappings of the Storage Index.
    /// </summary>
    public RevisionManifest GetRevision(ExtendedGuid revisionId)
    {
        foreach (var mapping in StorageIndex.Entries.OfType<StorageIndexRevisionMapping>())
        {
            if (mapping.Revision == revisionId)
            {
                return Get<RevisionManifest>(mapping.MappingId);
            }
        }

        throw new FssHttpDException($"No revision mapping for the revision {revisionId}");
    }

    /// <summary>
    /// The storage without the data elements that cannot be reached from the Storage Index: those of revisions no longer
    /// mapped, and the Object Groups and Object Data BLOBs only they referenced.
    /// </summary>
    public CellStorage Prune()
    {
        var reachable = new Dictionary<ExtendedGuid, DataElement>();
        foreach (var id in Reached())
        {
            if (Elements.TryGetValue(id, out var element))
            {
                reachable.TryAdd(id, element);
            }
        }

        return this with
        {
            Elements = reachable
        };
    }

    /// <summary>
    /// Whether the Storage Index reaches a data element the storage does not have.
    /// </summary>
    public bool HasMissing() =>
        Reached().Any(_ => !_.IsNull && !Elements.ContainsKey(_));

    /// <summary>
    /// The identifiers the Storage Index reaches, through the data elements the storage has: its mappings, and the Object
    /// Groups and Object Data BLOBs of the revisions it maps.
    /// </summary>
    IEnumerable<ExtendedGuid> Reached()
    {
        yield return StorageIndexId;
        foreach (var entry in StorageIndex.Entries)
        {
            yield return entry.MappingId;
            if (entry is not StorageIndexRevisionMapping ||
                !Elements.TryGetValue(entry.MappingId, out var element) ||
                element is not RevisionManifest revision)
            {
                continue;
            }

            foreach (var groupId in revision.ObjectGroups)
            {
                yield return groupId;
                if (!Elements.TryGetValue(groupId, out var groupElement) ||
                    groupElement is not ObjectGroup group)
                {
                    continue;
                }

                foreach (var declaration in group.Declarations.OfType<ObjectDataBlobDeclaration>())
                {
                    yield return declaration.BlobId;
                }

                foreach (var data in group.Data.OfType<ObjectDataBlobReference>())
                {
                    yield return data.BlobId;
                }
            }
        }
    }
}

static class DataElementDictionary
{
    public static Dictionary<ExtendedGuid, DataElement> With(this Dictionary<ExtendedGuid, DataElement> elements, IEnumerable<DataElement> added)
    {
        foreach (var element in added)
        {
            elements[element.Id] = element;
        }

        return elements;
    }
}
