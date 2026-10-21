/// <summary>
/// A Storage Manifest data element ([MS-FSSHTTPB] 2.2.1.12.3): the schema of the file and its roots.
/// </summary>
record StorageManifest(ExtendedGuid Id, SerialNumber SerialNumber, Guid SchemaId, IReadOnlyList<StorageManifestRoot> Roots) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.StorageManifest;

    public static StorageManifest ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
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

    protected override void WriteBody(ref FssHttpBWriter writer)
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