/// <summary>
/// A response ([MS-FSSHTTPB] 2.2.3): the binary content of the response to a Cell sub-request of [MS-FSSHTTP]. Either
/// the request failed as a whole, with an error, or each sub-request has its sub-response.
/// </summary>
record FssHttpBResponse(
    IReadOnlyList<SubResponse> SubResponses,
    DataElementPackage? DataElements = null,
    ResponseError? Error = null,
    ushort ProtocolVersion = 12,
    ushort MinimumVersion = 11)
{
    /// <summary>
    /// The signature of a response, 0x9B069439F329CF9D.
    /// </summary>
    public const ulong Signature = 0x9B069439F329CF9D;

    public static FssHttpBResponse Failed(ResponseError error) => new([], Error: error);

    public static FssHttpBResponse Read(ReadOnlyMemory<byte> data)
    {
        var reader = new FssHttpBReader(data);
        var protocolVersion = reader.ReadUInt16();
        var minimumVersion = reader.ReadUInt16();
        var at = reader.Offset;
        if (reader.ReadUInt64() != Signature)
        {
            throw new FssHttpBException("Not an FSSHTTPB response signature", at);
        }

        var fields = reader.ReadStart(StreamObjectType.Response);
        var status = fields.ReadByte();
        fields.EnsureEnd("a response start");
        FssHttpBResponse response;
        if ((status & 0x01) != 0)
        {
            response = Failed(ResponseError.Read(ref reader));
        }
        else
        {
            DataElementPackage? package = null;
            if (reader.PeekStart(StreamObjectType.DataElementPackage))
            {
                package = DataElementPackage.Read(ref reader);
            }

            var subResponses = new List<SubResponse>();
            while (reader.PeekStart(StreamObjectType.SubResponse))
            {
                subResponses.Add(SubResponse.Read(ref reader));
            }

            response = new(subResponses, package);
        }

        reader.ReadEnd(StreamObjectType.Response);
        reader.EnsureEnd("the response");
        return response with
        {
            ProtocolVersion = protocolVersion,
            MinimumVersion = minimumVersion
        };
    }

    public void Write(IBufferWriter<byte> output)
    {
        var writer = new FssHttpBWriter(output);
        writer.WriteUInt16(ProtocolVersion);
        writer.WriteUInt16(MinimumVersion);
        writer.WriteUInt64(Signature);
        writer.WriteStart(StreamObjectType.Response, 1);
        if (Error != null)
        {
            writer.WriteByte(1);
            Error.Write(ref writer);
        }
        else
        {
            writer.WriteByte(0);
            DataElements?.Write(ref writer);
            foreach (var subResponse in SubResponses)
            {
                subResponse.Write(ref writer);
            }
        }

        writer.WriteEnd(StreamObjectType.Response);
    }

    public byte[] ToArray()
    {
        var output = new ArrayBufferWriter<byte>();
        Write(output);
        return output.WrittenSpan.ToArray();
    }
}