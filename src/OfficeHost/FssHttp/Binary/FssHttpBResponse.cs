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

/// <summary>
/// The sub-response of a sub-request that failed (the status bit of 2.2.3.1), with its error.
/// </summary>
record FailedSubResponse(ulong RequestId, RequestType FailedType, ResponseError Error) :
    SubResponse(RequestId)
{
    public override RequestType Type => FailedType;

    protected override void WriteData(ref FssHttpBWriter writer) => Error.Write(ref writer);
}

/// <summary>
/// A Query Access sub-response ([MS-FSSHTTPB] 2.2.3.1.1): whether reads and writes will succeed, as errors, the HRESULT
/// zero for yes.
/// </summary>
record QueryAccessResponse(ulong RequestId, ResponseError ReadAccess, ResponseError WriteAccess) :
    SubResponse(RequestId)
{
    public override RequestType Type => RequestType.QueryAccess;

    public static QueryAccessResponse ReadData(ref FssHttpBReader reader, ulong requestId)
    {
        reader.ReadStart(StreamObjectType.ReadAccessResponse).EnsureEnd("a read access response start");
        var read = ResponseError.Read(ref reader);
        reader.ReadEnd(StreamObjectType.ReadAccessResponse);
        reader.ReadStart(StreamObjectType.WriteAccessResponse).EnsureEnd("a write access response start");
        var write = ResponseError.Read(ref reader);
        reader.ReadEnd(StreamObjectType.WriteAccessResponse);
        return new(requestId, read, write);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.ReadAccessResponse, 0);
        ReadAccess.Write(ref writer);
        writer.WriteEnd(StreamObjectType.ReadAccessResponse);
        writer.WriteStart(StreamObjectType.WriteAccessResponse, 0);
        WriteAccess.Write(ref writer);
        writer.WriteEnd(StreamObjectType.WriteAccessResponse);
    }
}

/// <summary>
/// The File Hash of a Query Changes sub-response: the type of the hash and the hash.
/// </summary>
record FileHash(ulong Type, ReadOnlyMemory<byte> Data);

/// <summary>
/// A Query Changes sub-response ([MS-FSSHTTPB] 2.2.3.1.2): the Storage Index of the file, and what the server knows of
/// it. The data elements are in the Data Element Package of the response.
/// </summary>
/// <param name="Flags">P, partial (0x01); A, user content equivalent version returned (0x02).</param>
record QueryChangesResponse(ulong RequestId, ExtendedGuid StorageIndex, byte Flags, Knowledge Knowledge, FileHash? FileHash = null) :
    SubResponse(RequestId)
{
    public override RequestType Type => RequestType.QueryChanges;

    public bool Partial => (Flags & 0x01) != 0;

    public static QueryChangesResponse ReadData(ref FssHttpBReader reader, ulong requestId)
    {
        var fields = reader.ReadStart(StreamObjectType.QueryChangesResponse);
        var storageIndex = fields.ReadExtendedGuid();
        var flags = fields.ReadByte();
        fields.EnsureEnd("a Query Changes response");
        var knowledge = Knowledge.Read(ref reader);
        FileHash? hash = null;
        if (reader.PeekStart(StreamObjectType.FileHash))
        {
            var hashFields = reader.ReadStart(StreamObjectType.FileHash);
            hash = new(hashFields.ReadCompactUInt64(), hashFields.ReadBinaryItem());
            hashFields.EnsureEnd("a file hash");
        }

        return new(requestId, storageIndex, flags, knowledge, hash);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.QueryChangesResponse, StorageIndex.Size + 1);
        writer.WriteExtendedGuid(StorageIndex);
        writer.WriteByte(Flags);
        Knowledge.Write(ref writer);
        if (FileHash != null)
        {
            writer.WriteStart(StreamObjectType.FileHash, CompactUInt64.Size(FileHash.Type) + FssHttpBSize.BinaryItem(FileHash.Data.Length));
            writer.WriteCompactUInt64(FileHash.Type);
            writer.WriteBinaryItem(FileHash.Data.Span);
        }
    }
}

/// <summary>
/// A Put Changes sub-response ([MS-FSSHTTPB] 2.2.3.1.3): the knowledge of the server after the changes. The Put Changes
/// response object, with the applied Storage Index and the data elements added, is read as optional: the example of
/// [MS-FSSHTTPB] 4.4 has none (docs/fsshttp-notes.md).
/// </summary>
record PutChangesResponse(
    ulong RequestId,
    Knowledge Knowledge,
    bool HasResponseObject = false,
    ExtendedGuid? AppliedStorageIndex = null,
    IReadOnlyList<ExtendedGuid>? DataElementsAdded = null,
    byte? DiagnosticOutput = null) :
    SubResponse(RequestId)
{
    public override RequestType Type => RequestType.PutChanges;

    public static PutChangesResponse ReadData(ref FssHttpBReader reader, ulong requestId)
    {
        var hasObject = false;
        ExtendedGuid? applied = null;
        ExtendedGuid[]? added = null;
        if (reader.PeekStart(StreamObjectType.PutChangesResponse))
        {
            hasObject = true;
            var fields = reader.ReadStart(StreamObjectType.PutChangesResponse);
            if (!fields.AtEnd)
            {
                applied = fields.ReadExtendedGuid();
            }

            if (!fields.AtEnd)
            {
                added = fields.ReadExtendedGuidArray();
            }

            fields.EnsureEnd("a Put Changes response");
        }

        var knowledge = Knowledge.Read(ref reader);
        byte? diagnostic = null;
        if (reader.PeekStart(StreamObjectType.DiagnosticRequestOptionOutput))
        {
            var fields = reader.ReadStart(StreamObjectType.DiagnosticRequestOptionOutput);
            diagnostic = fields.ReadByte();
            fields.EnsureEnd("a diagnostic request option output");
        }

        return new(requestId, knowledge, hasObject, applied, added, diagnostic);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        if (HasResponseObject || AppliedStorageIndex != null || DataElementsAdded != null)
        {
            var size = 0;
            if (AppliedStorageIndex != null || DataElementsAdded != null)
            {
                size += (AppliedStorageIndex ?? ExtendedGuid.Null).Size;
            }

            if (DataElementsAdded != null)
            {
                size += FssHttpBSize.ExtendedGuidArray(DataElementsAdded);
            }

            writer.WriteStart(StreamObjectType.PutChangesResponse, size);
            if (AppliedStorageIndex != null || DataElementsAdded != null)
            {
                writer.WriteExtendedGuid(AppliedStorageIndex ?? ExtendedGuid.Null);
            }

            if (DataElementsAdded != null)
            {
                writer.WriteExtendedGuidArray(DataElementsAdded);
            }
        }

        Knowledge.Write(ref writer);
        if (DiagnosticOutput != null)
        {
            writer.WriteStart(StreamObjectType.DiagnosticRequestOptionOutput, 1);
            writer.WriteByte(DiagnosticOutput.Value);
        }
    }
}

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
