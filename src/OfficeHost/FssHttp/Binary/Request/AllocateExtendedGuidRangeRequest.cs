/// <summary>
/// An Allocate Extended GUID Range sub-request ([MS-FSSHTTPB] 2.2.2.1.5): the number of Extended GUIDs to allocate.
/// </summary>
record AllocateExtendedGuidRangeRequest(ulong RequestId, ulong Count, ulong Priority = 0, Guid? TargetPartition = null) :
    SubRequest(RequestId, Priority, TargetPartition)
{
    public override RequestType Type => RequestType.AllocateExtendedGuidRange;

    public static AllocateExtendedGuidRangeRequest ReadData(ref FssHttpBReader reader, ulong requestId, ulong priority, Guid? partition)
    {
        var fields = reader.ReadStart(StreamObjectType.AllocateExtendedGuidRangeRequest);
        var count = fields.ReadCompactUInt64();
        // the reserved byte, which is ignored
        fields.ReadByte();
        fields.EnsureEnd("an Allocate Extended GUID Range request");
        return new(requestId, count, priority, partition);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.AllocateExtendedGuidRangeRequest, CompactUInt64.Size(Count) + 1);
        writer.WriteCompactUInt64(Count);
        writer.WriteByte(0);
    }
}