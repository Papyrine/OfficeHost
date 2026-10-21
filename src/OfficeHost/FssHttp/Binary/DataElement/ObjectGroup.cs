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

    public static ObjectGroup ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
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

    static ObjectGroupDeclaration ReadDeclaration(ref FssHttpBReader reader)
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

    static ObjectGroupData ReadData(ref FssHttpBReader reader)
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

    protected override void WriteBody(ref FssHttpBWriter writer)
    {
        if (Hash != null)
        {
            writer.WriteStart(StreamObjectType.DataElementHash, CompactUInt64.Size(Hash.Scheme) + FssHttpBSize.BinaryItem(Hash.Data.Length));
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

    static void WriteDeclaration(ref FssHttpBWriter writer, ObjectGroupDeclaration declaration)
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

    static void WriteData(ref FssHttpBWriter writer, ObjectGroupData data)
    {
        var references = FssHttpBSize.ExtendedGuidArray(data.ObjectReferences) + FssHttpBSize.CellIdArray(data.CellReferences);
        switch (data)
        {
            case ObjectData value:
                writer.WriteStart(StreamObjectType.ObjectGroupObjectData, references + FssHttpBSize.BinaryItem(value.Data.Length));
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