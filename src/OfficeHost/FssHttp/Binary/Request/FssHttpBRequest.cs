/// <summary>
/// A request ([MS-FSSHTTPB] 2.2.2): the binary content of a Cell sub-request of [MS-FSSHTTP]. The user agent is optional:
/// Word sends none with a lone Query Access sub-request (docs/fsshttp-notes.md).
/// </summary>
record FssHttpBRequest(
    UserAgent? UserAgent,
    IReadOnlyList<SubRequest> SubRequests,
    DataElementPackage? DataElements = null,
    RequestHashingOptions? Hashing = null,
    byte? CellRoundtripOptions = null,
    ushort ProtocolVersion = 12,
    ushort MinimumVersion = 11)
{
    /// <summary>
    /// The signature of a request, 0x9B069439F329CF9C.
    /// </summary>
    public const ulong Signature = 0x9B069439F329CF9C;

    public static FssHttpBRequest Read(ReadOnlyMemory<byte> data)
    {
        var reader = new FssHttpBReader(data);
        var protocolVersion = reader.ReadUInt16();
        var minimumVersion = reader.ReadUInt16();
        var at = reader.Offset;
        if (reader.ReadUInt64() != Signature)
        {
            throw new FssHttpBException("Not an FSSHTTPB request signature", at);
        }

        reader.ReadStart(StreamObjectType.Request).EnsureEnd("a request start");
        UserAgent? userAgent = null;
        if (reader.PeekStart(StreamObjectType.UserAgent))
        {
            userAgent = UserAgent.Read(ref reader);
        }

        RequestHashingOptions? hashing = null;
        if (reader.PeekStart(StreamObjectType.RequestHashingOptions))
        {
            var fields = reader.ReadStart(StreamObjectType.RequestHashingOptions);
            hashing = new(fields.ReadCompactUInt64(), fields.ReadByte());
            fields.EnsureEnd("the request hashing options");
        }

        byte? roundtrip = null;
        if (reader.PeekStart(StreamObjectType.CellRoundtripOptions))
        {
            var fields = reader.ReadStart(StreamObjectType.CellRoundtripOptions);
            roundtrip = fields.ReadByte();
            fields.EnsureEnd("the cell roundtrip options");
        }

        var subRequests = new List<SubRequest>();
        while (reader.PeekStart(StreamObjectType.SubRequest))
        {
            subRequests.Add(SubRequest.Read(ref reader));
        }

        DataElementPackage? package = null;
        if (reader.PeekStart(StreamObjectType.DataElementPackage))
        {
            package = DataElementPackage.Read(ref reader);
        }

        reader.ReadEnd(StreamObjectType.Request);
        reader.EnsureEnd("the request");
        return new(userAgent, subRequests, package, hashing, roundtrip, protocolVersion, minimumVersion);
    }

    public void Write(IBufferWriter<byte> output)
    {
        var writer = new FssHttpBWriter(output);
        writer.WriteUInt16(ProtocolVersion);
        writer.WriteUInt16(MinimumVersion);
        writer.WriteUInt64(Signature);
        writer.WriteStart(StreamObjectType.Request, 0);
        UserAgent?.Write(ref writer);
        if (Hashing != null)
        {
            writer.WriteStart(StreamObjectType.RequestHashingOptions, CompactUInt64.Size(Hashing.Schema) + 1);
            writer.WriteCompactUInt64(Hashing.Schema);
            writer.WriteByte(Hashing.Flags);
        }

        if (CellRoundtripOptions != null)
        {
            writer.WriteStart(StreamObjectType.CellRoundtripOptions, 1);
            writer.WriteByte(CellRoundtripOptions.Value);
        }

        foreach (var subRequest in SubRequests)
        {
            subRequest.Write(ref writer);
        }

        DataElements?.Write(ref writer);
        writer.WriteEnd(StreamObjectType.Request);
    }

    public byte[] ToArray()
    {
        var output = new ArrayBufferWriter<byte>();
        Write(output);
        return output.WrittenSpan.ToArray();
    }
}