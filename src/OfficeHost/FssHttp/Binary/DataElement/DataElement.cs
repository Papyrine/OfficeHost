/// <summary>
/// A data element ([MS-FSSHTTPB] 2.2.1.12.2 to 2.2.1.12.8): an identifier, a Serial Number and a type, then the data of
/// the type. Data elements are immutable.
/// </summary>
abstract record DataElement(ExtendedGuid Id, SerialNumber SerialNumber)
{
    public abstract DataElementType Type { get; }

    public static DataElement Read(ref FssHttpBReader reader)
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
            _ => throw new FssHttpBException($"Unknown data element type {type}", at)
        };
        reader.ReadEnd(StreamObjectType.DataElement);
        return element;
    }

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.DataElement, Id.Size + SerialNumber.Size + CompactUInt64.Size((ulong) Type));
        writer.WriteExtendedGuid(Id);
        writer.WriteSerialNumber(SerialNumber);
        writer.WriteCompactUInt64((ulong) Type);
        WriteBody(ref writer);
        writer.WriteEnd(StreamObjectType.DataElement);
    }

    protected abstract void WriteBody(ref FssHttpBWriter writer);
}