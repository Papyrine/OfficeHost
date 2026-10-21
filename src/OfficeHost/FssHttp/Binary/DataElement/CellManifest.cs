/// <summary>
/// A Cell Manifest data element ([MS-FSSHTTPB] 2.2.1.12.4): the current revision of a cell.
/// </summary>
record CellManifest(ExtendedGuid Id, SerialNumber SerialNumber, ExtendedGuid CurrentRevision) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.CellManifest;

    public static CellManifest ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.CellManifestCurrentRevision);
        var revision = fields.ReadExtendedGuid();
        fields.EnsureEnd("a Cell Manifest current revision");
        return new(id, serialNumber, revision);
    }

    protected override void WriteBody(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.CellManifestCurrentRevision, CurrentRevision.Size);
        writer.WriteExtendedGuid(CurrentRevision);
    }
}