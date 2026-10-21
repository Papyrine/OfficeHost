/// <summary>
/// A Storage Index data element ([MS-FSSHTTPB] 2.2.1.12.2): the mappings, in no set order, which is kept as it is.
/// </summary>
record StorageIndex(ExtendedGuid Id, SerialNumber SerialNumber, IReadOnlyList<StorageIndexEntry> Entries) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.StorageIndex;

    public static StorageIndex ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var entries = new List<StorageIndexEntry>();
        while (!reader.PeekEnd())
        {
            var header = reader.PeekHeader();
            var fields = reader.ReadStart(header.Type);
            StorageIndexEntry entry = header.Type switch
            {
                StreamObjectType.StorageIndexManifestMapping =>
                    new StorageIndexManifestMapping(fields.ReadExtendedGuid(), fields.ReadSerialNumber()),
                StreamObjectType.StorageIndexCellMapping =>
                    new StorageIndexCellMapping(fields.ReadCellId(), fields.ReadExtendedGuid(), fields.ReadSerialNumber()),
                StreamObjectType.StorageIndexRevisionMapping =>
                    new StorageIndexRevisionMapping(fields.ReadExtendedGuid(), fields.ReadExtendedGuid(), fields.ReadSerialNumber()),
                _ => throw fields.Error($"{header} in a Storage Index")
            };
            fields.EnsureEnd("a Storage Index mapping");
            entries.Add(entry);
        }

        return new(id, serialNumber, entries);
    }

    protected override void WriteBody(ref FssHttpBWriter writer)
    {
        foreach (var entry in Entries)
        {
            switch (entry)
            {
                case StorageIndexManifestMapping manifest:
                    writer.WriteStart(StreamObjectType.StorageIndexManifestMapping, manifest.MappingId.Size + manifest.MappingSerialNumber.Size);
                    break;
                case StorageIndexCellMapping cell:
                    writer.WriteStart(StreamObjectType.StorageIndexCellMapping, cell.Cell.Size + cell.MappingId.Size + cell.MappingSerialNumber.Size);
                    writer.WriteCellId(cell.Cell);
                    break;
                case StorageIndexRevisionMapping revision:
                    writer.WriteStart(StreamObjectType.StorageIndexRevisionMapping, revision.Revision.Size + revision.MappingId.Size + revision.MappingSerialNumber.Size);
                    writer.WriteExtendedGuid(revision.Revision);
                    break;
            }

            writer.WriteExtendedGuid(entry.MappingId);
            writer.WriteSerialNumber(entry.MappingSerialNumber);
        }
    }
}