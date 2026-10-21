/// <summary>
/// An Allocate Extended GUID Range sub-response ([MS-FSSHTTPB] 2.2.3.1.4): the GUID of the Extended GUIDs allocated, with
/// the values from Min to Max - 1.
/// </summary>
record AllocateExtendedGuidRangeResponse(ulong RequestId, Guid Guid, ulong Min, ulong Max) :
    SubResponse(RequestId)
{
    public override RequestType Type => RequestType.AllocateExtendedGuidRange;

    public static AllocateExtendedGuidRangeResponse ReadData(ref FssHttpBReader reader, ulong requestId)
    {
        var fields = reader.ReadStart(StreamObjectType.AllocateExtendedGuidRangeResponse);
        var guid = fields.ReadGuid();
        var min = fields.ReadCompactUInt64();
        var max = fields.ReadCompactUInt64();
        fields.EnsureEnd("an Allocate Extended GUID Range response");
        return new(requestId, guid, min, max);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.AllocateExtendedGuidRangeResponse, 16 + CompactUInt64.Size(Min) + CompactUInt64.Size(Max));
        writer.WriteGuid(Guid);
        writer.WriteCompactUInt64(Min);
        writer.WriteCompactUInt64(Max);
    }
}