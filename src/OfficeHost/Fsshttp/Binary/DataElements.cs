/// <summary>
/// The data element types of [MS-FSSHTTPB] 2.2.1.12.1.
/// </summary>
enum DataElementType : ulong
{
    StorageIndex = 1,
    StorageManifest = 2,
    CellManifest = 3,
    RevisionManifest = 4,
    ObjectGroup = 5,
    DataElementFragment = 6,
    ObjectDataBlob = 10
}

/// <summary>
/// A Data Element Package ([MS-FSSHTTPB] 2.2.1.12): the data elements of a request or a response.
/// </summary>
record DataElementPackage(IReadOnlyList<DataElement> Elements)
{
    public static DataElementPackage Read(ref FsshttpbReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.DataElementPackage);
        // the Reserved byte, which is ignored
        fields.ReadByte();
        fields.EnsureEnd("the Data Element Package start");
        var elements = new List<DataElement>();
        while (reader.PeekStart(StreamObjectType.DataElement))
        {
            elements.Add(DataElement.Read(ref reader));
        }

        reader.ReadEnd(StreamObjectType.DataElementPackage);
        return new(elements);
    }

    public void Write(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.DataElementPackage, 1);
        writer.WriteByte(0);
        foreach (var element in Elements)
        {
            element.Write(ref writer);
        }

        writer.WriteEnd(StreamObjectType.DataElementPackage);
    }
}

/// <summary>
/// A data element ([MS-FSSHTTPB] 2.2.1.12.2 to 2.2.1.12.8): an identifier, a Serial Number and a type, then the data of
/// the type. Data elements are immutable.
/// </summary>
abstract record DataElement(ExtendedGuid Id, SerialNumber SerialNumber)
{
    public abstract DataElementType Type { get; }

    public static DataElement Read(ref FsshttpbReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.DataElement);
        var at = fields.Offset;
        var id = fields.ReadExtendedGuid();
        var serialNumber = fields.ReadSerialNumber();
        var type = fields.ReadCompactUInt64();
        fields.EnsureEnd("the data element start");
        DataElement element = (DataElementType) type switch
        {
            DataElementType.StorageIndex => StorageIndex.ReadBody(ref reader, id, serialNumber),
            DataElementType.StorageManifest => StorageManifest.ReadBody(ref reader, id, serialNumber),
            DataElementType.CellManifest => CellManifest.ReadBody(ref reader, id, serialNumber),
            DataElementType.RevisionManifest => RevisionManifest.ReadBody(ref reader, id, serialNumber),
            DataElementType.ObjectGroup => ObjectGroup.ReadBody(ref reader, id, serialNumber),
            DataElementType.DataElementFragment => DataElementFragment.ReadBody(ref reader, id, serialNumber),
            DataElementType.ObjectDataBlob => ObjectDataBlob.ReadBody(ref reader, id, serialNumber),
            _ => throw new FsshttpbException($"Unknown data element type {type}", at)
        };
        reader.ReadEnd(StreamObjectType.DataElement);
        return element;
    }

    public void Write(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.DataElement, Id.Size + SerialNumber.Size + CompactUInt64.Size((ulong) Type));
        writer.WriteExtendedGuid(Id);
        writer.WriteSerialNumber(SerialNumber);
        writer.WriteCompactUInt64((ulong) Type);
        WriteBody(ref writer);
        writer.WriteEnd(StreamObjectType.DataElement);
    }

    protected abstract void WriteBody(ref FsshttpbWriter writer);
}

/// <summary>
/// A mapping of a Storage Index ([MS-FSSHTTPB] 2.2.1.12.2), each with the Serial Number of the mapping.
/// </summary>
abstract record StorageIndexEntry(ExtendedGuid MappingId, SerialNumber MappingSerialNumber);

/// <summary>
/// The Storage Index Manifest Mapping: the Storage Manifest.
/// </summary>
record StorageIndexManifestMapping(ExtendedGuid MappingId, SerialNumber MappingSerialNumber) :
    StorageIndexEntry(MappingId, MappingSerialNumber);

/// <summary>
/// A Storage Index Cell Mapping: the Cell Manifest of a cell.
/// </summary>
record StorageIndexCellMapping(CellId Cell, ExtendedGuid MappingId, SerialNumber MappingSerialNumber) :
    StorageIndexEntry(MappingId, MappingSerialNumber);

/// <summary>
/// A Storage Index Revision Mapping: the Revision Manifest of a revision.
/// </summary>
record StorageIndexRevisionMapping(ExtendedGuid Revision, ExtendedGuid MappingId, SerialNumber MappingSerialNumber) :
    StorageIndexEntry(MappingId, MappingSerialNumber);

/// <summary>
/// A Storage Index data element ([MS-FSSHTTPB] 2.2.1.12.2): the mappings, in no set order, which is kept as it is.
/// </summary>
record StorageIndex(ExtendedGuid Id, SerialNumber SerialNumber, IReadOnlyList<StorageIndexEntry> Entries) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.StorageIndex;

    public static StorageIndex ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
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

    protected override void WriteBody(ref FsshttpbWriter writer)
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

/// <summary>
/// A root of a Storage Manifest: a root identifier and its cell.
/// </summary>
record StorageManifestRoot(ExtendedGuid RootId, CellId Cell);

/// <summary>
/// A Storage Manifest data element ([MS-FSSHTTPB] 2.2.1.12.3): the schema of the file and its roots.
/// </summary>
record StorageManifest(ExtendedGuid Id, SerialNumber SerialNumber, Guid SchemaId, IReadOnlyList<StorageManifestRoot> Roots) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.StorageManifest;

    public static StorageManifest ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var schema = reader.ReadStart(StreamObjectType.StorageManifestSchemaGuid);
        var schemaId = schema.ReadGuid();
        schema.EnsureEnd("a Storage Manifest schema GUID");
        var roots = new List<StorageManifestRoot>();
        while (reader.PeekStart(StreamObjectType.StorageManifestRootDeclare))
        {
            var root = reader.ReadStart(StreamObjectType.StorageManifestRootDeclare);
            roots.Add(new(root.ReadExtendedGuid(), root.ReadCellId()));
            root.EnsureEnd("a Storage Manifest root declare");
        }

        return new(id, serialNumber, schemaId, roots);
    }

    protected override void WriteBody(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.StorageManifestSchemaGuid, 16);
        writer.WriteGuid(SchemaId);
        foreach (var root in Roots)
        {
            writer.WriteStart(StreamObjectType.StorageManifestRootDeclare, root.RootId.Size + root.Cell.Size);
            writer.WriteExtendedGuid(root.RootId);
            writer.WriteCellId(root.Cell);
        }
    }
}

/// <summary>
/// A Cell Manifest data element ([MS-FSSHTTPB] 2.2.1.12.4): the current revision of a cell.
/// </summary>
record CellManifest(ExtendedGuid Id, SerialNumber SerialNumber, ExtendedGuid CurrentRevision) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.CellManifest;

    public static CellManifest ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.CellManifestCurrentRevision);
        var revision = fields.ReadExtendedGuid();
        fields.EnsureEnd("a Cell Manifest current revision");
        return new(id, serialNumber, revision);
    }

    protected override void WriteBody(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.CellManifestCurrentRevision, CurrentRevision.Size);
        writer.WriteExtendedGuid(CurrentRevision);
    }
}

/// <summary>
/// A root of a revision: a root identifier and its object.
/// </summary>
record RevisionManifestRoot(ExtendedGuid RootId, ExtendedGuid ObjectId);

/// <summary>
/// A Revision Manifest data element ([MS-FSSHTTPB] 2.2.1.12.5): a revision, the revision it is based on, its roots and
/// its Object Groups.
/// </summary>
record RevisionManifest(
    ExtendedGuid Id,
    SerialNumber SerialNumber,
    ExtendedGuid RevisionId,
    ExtendedGuid BaseRevisionId,
    IReadOnlyList<RevisionManifestRoot> Roots,
    IReadOnlyList<ExtendedGuid> ObjectGroups) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.RevisionManifest;

    public static RevisionManifest ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.RevisionManifest);
        var revisionId = fields.ReadExtendedGuid();
        var baseRevisionId = fields.ReadExtendedGuid();
        fields.EnsureEnd("a Revision Manifest");
        var roots = new List<RevisionManifestRoot>();
        while (reader.PeekStart(StreamObjectType.RevisionManifestRootDeclare))
        {
            var root = reader.ReadStart(StreamObjectType.RevisionManifestRootDeclare);
            roots.Add(new(root.ReadExtendedGuid(), root.ReadExtendedGuid()));
            root.EnsureEnd("a Revision Manifest root declare");
        }

        var groups = new List<ExtendedGuid>();
        while (reader.PeekStart(StreamObjectType.RevisionManifestObjectGroupReferences))
        {
            var group = reader.ReadStart(StreamObjectType.RevisionManifestObjectGroupReferences);
            groups.Add(group.ReadExtendedGuid());
            group.EnsureEnd("a Revision Manifest Object Group reference");
        }

        return new(id, serialNumber, revisionId, baseRevisionId, roots, groups);
    }

    protected override void WriteBody(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.RevisionManifest, RevisionId.Size + BaseRevisionId.Size);
        writer.WriteExtendedGuid(RevisionId);
        writer.WriteExtendedGuid(BaseRevisionId);
        foreach (var root in Roots)
        {
            writer.WriteStart(StreamObjectType.RevisionManifestRootDeclare, root.RootId.Size + root.ObjectId.Size);
            writer.WriteExtendedGuid(root.RootId);
            writer.WriteExtendedGuid(root.ObjectId);
        }

        foreach (var group in ObjectGroups)
        {
            writer.WriteStart(StreamObjectType.RevisionManifestObjectGroupReferences, group.Size);
            writer.WriteExtendedGuid(group);
        }
    }
}

/// <summary>
/// A Data Element Hash ([MS-FSSHTTPB] 2.2.1.12.6.6): the hash scheme (1, [MS-PCCRC] 2.3) and the hash.
/// </summary>
record DataElementHash(ulong Scheme, ReadOnlyMemory<byte> Data);

/// <summary>
/// A declaration of an Object Group: an Object Declaration (2.2.1.12.6.1) or an Object Data BLOB Declaration
/// (2.2.1.12.6.2).
/// </summary>
abstract record ObjectGroupDeclaration(ExtendedGuid ObjectId, ulong PartitionId, ulong ObjectReferencesCount, ulong CellReferencesCount);

/// <summary>
/// An Object Declaration ([MS-FSSHTTPB] 2.2.1.12.6.1).
/// </summary>
record ObjectDeclaration(ExtendedGuid ObjectId, ulong PartitionId, ulong DataSize, ulong ObjectReferencesCount, ulong CellReferencesCount) :
    ObjectGroupDeclaration(ObjectId, PartitionId, ObjectReferencesCount, CellReferencesCount);

/// <summary>
/// An Object Data BLOB Declaration ([MS-FSSHTTPB] 2.2.1.12.6.2): an object whose data is in an Object Data BLOB.
/// </summary>
record ObjectDataBlobDeclaration(ExtendedGuid ObjectId, ExtendedGuid BlobId, ulong PartitionId, ulong ObjectReferencesCount, ulong CellReferencesCount) :
    ObjectGroupDeclaration(ObjectId, PartitionId, ObjectReferencesCount, CellReferencesCount);

/// <summary>
/// The data of an object of an Object Group: Object Data (2.2.1.12.6.4) with its data, Object Excluded Data with the
/// size of the data only, or an Object Data BLOB Reference (2.2.1.12.6.5).
/// </summary>
abstract record ObjectGroupData(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences);

record ObjectData(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences, ReadOnlyMemory<byte> Data) :
    ObjectGroupData(ObjectReferences, CellReferences);

record ObjectExcludedData(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences, ulong DataSize) :
    ObjectGroupData(ObjectReferences, CellReferences);

record ObjectDataBlobReference(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences, ExtendedGuid BlobId) :
    ObjectGroupData(ObjectReferences, CellReferences);

/// <summary>
/// An Object Group data element ([MS-FSSHTTPB] 2.2.1.12.6): an optional hash, the declarations of the objects, their
/// optional metadata (the change frequencies of 2.2.1.12.6.3.1), and their data.
/// </summary>
record ObjectGroup(
    ExtendedGuid Id,
    SerialNumber SerialNumber,
    DataElementHash? Hash,
    IReadOnlyList<ObjectGroupDeclaration> Declarations,
    IReadOnlyList<ulong>? Metadata,
    IReadOnlyList<ObjectGroupData> Data) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.ObjectGroup;

    public static ObjectGroup ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        DataElementHash? hash = null;
        if (reader.PeekStart(StreamObjectType.DataElementHash))
        {
            var fields = reader.ReadStart(StreamObjectType.DataElementHash);
            hash = new(fields.ReadCompactUInt64(), fields.ReadBinaryItem());
            fields.EnsureEnd("a Data Element Hash");
        }

        var declarationsStart = reader.ReadStart(StreamObjectType.ObjectGroupDeclarations);
        declarationsStart.EnsureEnd("an Object Group declarations start");
        var declarations = new List<ObjectGroupDeclaration>();
        while (!reader.PeekEnd())
        {
            declarations.Add(ReadDeclaration(ref reader));
        }

        reader.ReadEnd(StreamObjectType.ObjectGroupDeclarations);

        List<ulong>? metadata = null;
        if (reader.PeekStart(StreamObjectType.ObjectGroupMetadataDeclarations))
        {
            var metadataStart = reader.ReadStart(StreamObjectType.ObjectGroupMetadataDeclarations);
            metadataStart.EnsureEnd("an Object Group metadata declarations start");
            metadata = [];
            while (reader.PeekStart(StreamObjectType.ObjectGroupMetadata))
            {
                var fields = reader.ReadStart(StreamObjectType.ObjectGroupMetadata);
                metadata.Add(fields.ReadCompactUInt64());
                fields.EnsureEnd("an Object Group metadata");
            }

            reader.ReadEnd(StreamObjectType.ObjectGroupMetadataDeclarations);
        }

        var dataStart = reader.ReadStart(StreamObjectType.ObjectGroupData);
        dataStart.EnsureEnd("an Object Group data start");
        var data = new List<ObjectGroupData>();
        while (!reader.PeekEnd())
        {
            data.Add(ReadData(ref reader));
        }

        reader.ReadEnd(StreamObjectType.ObjectGroupData);
        return new(id, serialNumber, hash, declarations, metadata, data);
    }

    static ObjectGroupDeclaration ReadDeclaration(ref FsshttpbReader reader)
    {
        var header = reader.PeekHeader();
        var fields = reader.ReadStart(header.Type);
        ObjectGroupDeclaration declaration = header.Type switch
        {
            StreamObjectType.ObjectGroupObjectDeclare =>
                new ObjectDeclaration(fields.ReadExtendedGuid(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64()),
            StreamObjectType.ObjectGroupObjectDataBlobDeclaration =>
                new ObjectDataBlobDeclaration(fields.ReadExtendedGuid(), fields.ReadExtendedGuid(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64()),
            _ => throw fields.Error($"{header} in Object Group declarations")
        };
        fields.EnsureEnd("an Object Group declaration");
        return declaration;
    }

    static ObjectGroupData ReadData(ref FsshttpbReader reader)
    {
        var header = reader.PeekHeader();
        var fields = reader.ReadStart(header.Type);
        ObjectGroupData data = header.Type switch
        {
            StreamObjectType.ObjectGroupObjectData =>
                new ObjectData(fields.ReadExtendedGuidArray(), fields.ReadCellIdArray(), fields.ReadBinaryItem()),
            StreamObjectType.ObjectGroupObjectExcludedData =>
                new ObjectExcludedData(fields.ReadExtendedGuidArray(), fields.ReadCellIdArray(), fields.ReadCompactUInt64()),
            StreamObjectType.ObjectGroupObjectDataBlobReference =>
                new ObjectDataBlobReference(fields.ReadExtendedGuidArray(), fields.ReadCellIdArray(), fields.ReadExtendedGuid()),
            _ => throw fields.Error($"{header} in Object Group data")
        };
        fields.EnsureEnd("an Object Group data");
        return data;
    }

    protected override void WriteBody(ref FsshttpbWriter writer)
    {
        if (Hash != null)
        {
            writer.WriteStart(StreamObjectType.DataElementHash, CompactUInt64.Size(Hash.Scheme) + FsshttpbSize.BinaryItem(Hash.Data.Length));
            writer.WriteCompactUInt64(Hash.Scheme);
            writer.WriteBinaryItem(Hash.Data.Span);
        }

        writer.WriteStart(StreamObjectType.ObjectGroupDeclarations, 0);
        foreach (var declaration in Declarations)
        {
            WriteDeclaration(ref writer, declaration);
        }

        writer.WriteEnd(StreamObjectType.ObjectGroupDeclarations);

        if (Metadata != null)
        {
            writer.WriteStart(StreamObjectType.ObjectGroupMetadataDeclarations, 0);
            foreach (var frequency in Metadata)
            {
                writer.WriteStart(StreamObjectType.ObjectGroupMetadata, CompactUInt64.Size(frequency));
                writer.WriteCompactUInt64(frequency);
            }

            writer.WriteEnd(StreamObjectType.ObjectGroupMetadataDeclarations);
        }

        writer.WriteStart(StreamObjectType.ObjectGroupData, 0);
        foreach (var data in Data)
        {
            WriteData(ref writer, data);
        }

        writer.WriteEnd(StreamObjectType.ObjectGroupData);
    }

    static void WriteDeclaration(ref FsshttpbWriter writer, ObjectGroupDeclaration declaration)
    {
        var counts = CompactUInt64.Size(declaration.PartitionId) +
                     CompactUInt64.Size(declaration.ObjectReferencesCount) +
                     CompactUInt64.Size(declaration.CellReferencesCount);
        switch (declaration)
        {
            case ObjectDeclaration value:
                writer.WriteStart(StreamObjectType.ObjectGroupObjectDeclare, value.ObjectId.Size + counts + CompactUInt64.Size(value.DataSize));
                writer.WriteExtendedGuid(value.ObjectId);
                writer.WriteCompactUInt64(value.PartitionId);
                writer.WriteCompactUInt64(value.DataSize);
                break;
            case ObjectDataBlobDeclaration value:
                writer.WriteStart(StreamObjectType.ObjectGroupObjectDataBlobDeclaration, value.ObjectId.Size + value.BlobId.Size + counts);
                writer.WriteExtendedGuid(value.ObjectId);
                writer.WriteExtendedGuid(value.BlobId);
                writer.WriteCompactUInt64(value.PartitionId);
                break;
        }

        writer.WriteCompactUInt64(declaration.ObjectReferencesCount);
        writer.WriteCompactUInt64(declaration.CellReferencesCount);
    }

    static void WriteData(ref FsshttpbWriter writer, ObjectGroupData data)
    {
        var references = FsshttpbSize.ExtendedGuidArray(data.ObjectReferences) + FsshttpbSize.CellIdArray(data.CellReferences);
        switch (data)
        {
            case ObjectData value:
                writer.WriteStart(StreamObjectType.ObjectGroupObjectData, references + FsshttpbSize.BinaryItem(value.Data.Length));
                break;
            case ObjectExcludedData value:
                writer.WriteStart(StreamObjectType.ObjectGroupObjectExcludedData, references + CompactUInt64.Size(value.DataSize));
                break;
            case ObjectDataBlobReference value:
                writer.WriteStart(StreamObjectType.ObjectGroupObjectDataBlobReference, references + value.BlobId.Size);
                break;
        }

        writer.WriteExtendedGuidArray(data.ObjectReferences);
        writer.WriteCellIdArray(data.CellReferences);
        switch (data)
        {
            case ObjectData value:
                writer.WriteBinaryItem(value.Data.Span);
                break;
            case ObjectExcludedData value:
                writer.WriteCompactUInt64(value.DataSize);
                break;
            case ObjectDataBlobReference value:
                writer.WriteExtendedGuid(value.BlobId);
                break;
        }
    }
}

/// <summary>
/// A Data Element Fragment data element ([MS-FSSHTTPB] 2.2.1.12.7): a part of a large data element.
/// </summary>
record DataElementFragment(
    ExtendedGuid Id,
    SerialNumber SerialNumber,
    ExtendedGuid FragmentId,
    ulong DataElementSize,
    FileChunkReference Chunk,
    ReadOnlyMemory<byte> Data) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.DataElementFragment;

    public static DataElementFragment ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.DataElementFragment);
        return new(id, serialNumber, fields.ReadExtendedGuid(), fields.ReadCompactUInt64(), fields.ReadFileChunkReference(), fields.ReadRest());
    }

    protected override void WriteBody(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.DataElementFragment, FragmentId.Size + CompactUInt64.Size(DataElementSize) + Chunk.Size + Data.Length);
        writer.WriteExtendedGuid(FragmentId);
        writer.WriteCompactUInt64(DataElementSize);
        writer.WriteFileChunkReference(Chunk);
        writer.WriteBytes(Data.Span);
    }
}

/// <summary>
/// An Object Data BLOB data element ([MS-FSSHTTPB] 2.2.1.12.8): the data of objects declared by Object Data BLOB
/// Declarations.
/// </summary>
record ObjectDataBlob(ExtendedGuid Id, SerialNumber SerialNumber, ReadOnlyMemory<byte> Data) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.ObjectDataBlob;

    public static ObjectDataBlob ReadBody(ref FsshttpbReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.ObjectDataBlob);
        return new(id, serialNumber, fields.ReadRest());
    }

    protected override void WriteBody(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.ObjectDataBlob, Data.Length);
        writer.WriteBytes(Data.Span);
    }
}
