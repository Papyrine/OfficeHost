/// <summary>
/// A sub-response ([MS-FSSHTTPB] 2.2.3.1): the id and type of its sub-request, then either an error, or the data of the
/// type.
/// </summary>
abstract record SubResponse(ulong RequestId)
{
    public abstract RequestType Type { get; }

    public static SubResponse Read(ref FssHttpBReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.SubResponse);
        var at = fields.Offset;
        var requestId = fields.ReadCompactUInt64();
        var type = fields.ReadCompactUInt64();
        var status = fields.ReadByte();
        fields.EnsureEnd("a sub-response start");
        if (!Enum.IsDefined((RequestType) type))
        {
            throw new FssHttpBException($"Unknown request type {type}", at);
        }

        SubResponse subResponse;
        if ((status & 0x01) != 0)
        {
            subResponse = new FailedSubResponse(requestId, (RequestType) type, ResponseError.Read(ref reader));
        }
        else
        {
            subResponse = (RequestType) type switch
            {
                RequestType.QueryAccess => QueryAccessResponse.ReadData(ref reader, requestId),
                RequestType.QueryChanges => QueryChangesResponse.ReadData(ref reader, requestId),
                RequestType.PutChanges => PutChangesResponse.ReadData(ref reader, requestId),
                _ => AllocateExtendedGuidRangeResponse.ReadData(ref reader, requestId)
            };
        }

        reader.ReadEnd(StreamObjectType.SubResponse);
        return subResponse;
    }

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.SubResponse, CompactUInt64.Size(RequestId) + CompactUInt64.Size((ulong) Type) + 1);
        writer.WriteCompactUInt64(RequestId);
        writer.WriteCompactUInt64((ulong) Type);
        if (this is FailedSubResponse)
        {
            writer.WriteByte(1);
        }
        else
        {
            writer.WriteByte(0);
        }

        WriteData(ref writer);
        writer.WriteEnd(StreamObjectType.SubResponse);
    }

    protected abstract void WriteData(ref FssHttpBWriter writer);

    public byte[] ToArray()
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        Write(ref writer);
        return output.WrittenSpan.ToArray();
    }
}