/// <summary>
/// A sub-request ([MS-FSSHTTPB] 2.2.2.1): its id, which the sub-response repeats, its priority, an optional target
/// partition (2.2.2.1.1), and the data of its type.
/// </summary>
abstract record SubRequest(ulong RequestId, ulong Priority, Guid? TargetPartition)
{
    public abstract RequestType Type { get; }

    public static SubRequest Read(ref FssHttpBReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.SubRequest);
        var at = fields.Offset;
        var requestId = fields.ReadCompactUInt64();
        var type = fields.ReadCompactUInt64();
        var priority = fields.ReadCompactUInt64();
        fields.EnsureEnd("a sub-request start");
        Guid? partition = null;
        if (reader.PeekStart(StreamObjectType.TargetPartitionId))
        {
            var partitionFields = reader.ReadStart(StreamObjectType.TargetPartitionId);
            partition = partitionFields.ReadGuid();
            partitionFields.EnsureEnd("a target partition id");
        }

        SubRequest subRequest = (RequestType) type switch
        {
            RequestType.QueryAccess => new QueryAccessRequest(requestId, priority, partition),
            RequestType.QueryChanges => QueryChangesRequest.ReadData(ref reader, requestId, priority, partition),
            RequestType.PutChanges => PutChangesRequest.ReadData(ref reader, requestId, priority, partition),
            RequestType.AllocateExtendedGuidRange => AllocateExtendedGuidRangeRequest.ReadData(ref reader, requestId, priority, partition),
            _ => throw new FssHttpBException($"Unknown request type {type}", at)
        };
        reader.ReadEnd(StreamObjectType.SubRequest);
        return subRequest;
    }

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.SubRequest, CompactUInt64.Size(RequestId) + CompactUInt64.Size((ulong) Type) + CompactUInt64.Size(Priority));
        writer.WriteCompactUInt64(RequestId);
        writer.WriteCompactUInt64((ulong) Type);
        writer.WriteCompactUInt64(Priority);
        if (TargetPartition != null)
        {
            writer.WriteStart(StreamObjectType.TargetPartitionId, 16);
            writer.WriteGuid(TargetPartition.Value);
        }

        WriteData(ref writer);
        writer.WriteEnd(StreamObjectType.SubRequest);
    }

    protected abstract void WriteData(ref FssHttpBWriter writer);
}