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

    public static DataElementFragment ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.DataElementFragment);
        return new(id, serialNumber, fields.ReadExtendedGuid(), fields.ReadCompactUInt64(), fields.ReadFileChunkReference(), fields.ReadRest());
    }

    protected override void WriteBody(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.DataElementFragment, FragmentId.Size + CompactUInt64.Size(DataElementSize) + Chunk.Size + Data.Length);
        writer.WriteExtendedGuid(FragmentId);
        writer.WriteCompactUInt64(DataElementSize);
        writer.WriteFileChunkReference(Chunk);
        writer.WriteBytes(Data.Span);
    }
}