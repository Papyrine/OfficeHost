/// <summary>
/// An Object Data BLOB data element ([MS-FSSHTTPB] 2.2.1.12.8): the data of objects declared by Object Data BLOB
/// Declarations.
/// </summary>
record ObjectDataBlob(ExtendedGuid Id, SerialNumber SerialNumber, ReadOnlyMemory<byte> Data) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.ObjectDataBlob;

    public static ObjectDataBlob ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.ObjectDataBlob);
        return new(id, serialNumber, fields.ReadRest());
    }

    protected override void WriteBody(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.ObjectDataBlob, Data.Length);
        writer.WriteBytes(Data.Span);
    }
}