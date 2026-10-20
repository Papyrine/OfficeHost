using System.Text;

/// <summary>
/// The request types of [MS-FSSHTTPB] 2.2.1.6.
/// </summary>
enum RequestType : ulong
{
    QueryAccess = 1,
    QueryChanges = 2,
    PutChanges = 5,
    AllocateExtendedGuidRange = 11
}

/// <summary>
/// The user agent of a request ([MS-FSSHTTPB] 2.2.2): a GUID, a client and platform (which replace the GUID, though
/// Word sends both), and a version.
/// </summary>
record UserAgent(Guid? Guid, string? Client, string? Platform, uint Version)
{
    public static UserAgent Read(ref FsshttpbReader reader)
    {
        reader.ReadStart(StreamObjectType.UserAgent).EnsureEnd("a user agent start");
        Guid? guid = null;
        if (reader.PeekStart(StreamObjectType.UserAgentGuid))
        {
            var fields = reader.ReadStart(StreamObjectType.UserAgentGuid);
            guid = fields.ReadGuid();
            fields.EnsureEnd("a user agent GUID");
        }

        string? client = null;
        string? platform = null;
        if (reader.PeekStart(StreamObjectType.UserAgentClientAndPlatform))
        {
            var fields = reader.ReadStart(StreamObjectType.UserAgentClientAndPlatform);
            client = Encoding.UTF8.GetString(fields.ReadBinaryItem().Span);
            platform = Encoding.UTF8.GetString(fields.ReadBinaryItem().Span);
            fields.EnsureEnd("a user agent client and platform");
        }

        var version = reader.ReadStart(StreamObjectType.UserAgentVersion);
        var value = version.ReadUInt32();
        version.EnsureEnd("a user agent version");
        reader.ReadEnd(StreamObjectType.UserAgent);
        return new(guid, client, platform, value);
    }

    public void Write(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.UserAgent, 0);
        if (Guid != null)
        {
            writer.WriteStart(StreamObjectType.UserAgentGuid, 16);
            writer.WriteGuid(Guid.Value);
        }

        if (Client != null || Platform != null)
        {
            var client = Encoding.UTF8.GetBytes(Client ?? "");
            var platform = Encoding.UTF8.GetBytes(Platform ?? "");
            writer.WriteStart(StreamObjectType.UserAgentClientAndPlatform, FsshttpbSize.BinaryItem(client.Length) + FsshttpbSize.BinaryItem(platform.Length));
            writer.WriteBinaryItem(client);
            writer.WriteBinaryItem(platform);
        }

        writer.WriteStart(StreamObjectType.UserAgentVersion, 4);
        writer.WriteUInt32(Version);
        writer.WriteEnd(StreamObjectType.UserAgent);
    }
}

/// <summary>
/// The Request Hashing Options of a request ([MS-FSSHTTPB] 2.2.2): the hashing schema, and a byte of flags (C, request
/// data element hashes instead of data, 0x04; D, request data element hashes, 0x08).
/// </summary>
record RequestHashingOptions(ulong Schema, byte Flags);

/// <summary>
/// A request ([MS-FSSHTTPB] 2.2.2): the binary content of a Cell sub-request of [MS-FSSHTTP]. The user agent is optional:
/// Word sends none with a lone Query Access sub-request (docs/fsshttp-notes.md).
/// </summary>
record FsshttpbRequest(
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

    public static FsshttpbRequest Read(ReadOnlyMemory<byte> data)
    {
        var reader = new FsshttpbReader(data);
        var protocolVersion = reader.ReadUInt16();
        var minimumVersion = reader.ReadUInt16();
        var at = reader.Offset;
        if (reader.ReadUInt64() != Signature)
        {
            throw new FsshttpbException("Not an FSSHTTPB request signature", at);
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
        var writer = new FsshttpbWriter(output);
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

/// <summary>
/// A sub-request ([MS-FSSHTTPB] 2.2.2.1): its id, which the sub-response repeats, its priority, an optional target
/// partition (2.2.2.1.1), and the data of its type.
/// </summary>
abstract record SubRequest(ulong RequestId, ulong Priority, Guid? TargetPartition)
{
    public abstract RequestType Type { get; }

    public static SubRequest Read(ref FsshttpbReader reader)
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
            _ => throw new FsshttpbException($"Unknown request type {type}", at)
        };
        reader.ReadEnd(StreamObjectType.SubRequest);
        return subRequest;
    }

    public void Write(ref FsshttpbWriter writer)
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

    protected abstract void WriteData(ref FsshttpbWriter writer);
}

/// <summary>
/// A Query Access sub-request ([MS-FSSHTTPB] 2.2.2.1.2), which has no data.
/// </summary>
record QueryAccessRequest(ulong RequestId, ulong Priority = 0, Guid? TargetPartition = null) :
    SubRequest(RequestId, Priority, TargetPartition)
{
    public override RequestType Type => RequestType.QueryAccess;

    protected override void WriteData(ref FsshttpbWriter writer)
    {
    }
}

/// <summary>
/// The Query Changes Request Arguments ([MS-FSSHTTPB] 2.2.2.1.3): a byte of flags (F, include the Storage Manifest,
/// 0x01; G, include cell changes, 0x02), and the cell the query is scoped to, the null Cell ID for none.
/// </summary>
record QueryChangesArguments(byte Flags, CellId? Cell)
{
    public bool IncludeStorageManifest => (Flags & 0x01) != 0;

    public bool IncludeCellChanges => (Flags & 0x02) != 0;
}

/// <summary>
/// A filter of a Query Changes sub-request ([MS-FSSHTTPB] 2.2.2.1.3.1): its type (1 all, 2 data element type, 3
/// Storage Index referenced, 4 Cell ID, 5 custom, 6 data element IDs, 7 hierarchy), whether it includes (1) or
/// excludes (0), its data as the stream object of the type, and its optional flags.
/// </summary>
record QueryChangesFilter(byte FilterType, byte Operation, RawStreamObject? Data, byte? Flags);

/// <summary>
/// A Query Changes sub-request ([MS-FSSHTTPB] 2.2.2.1.3): what the client knows of the file, to receive what it does not.
/// </summary>
/// <param name="Flags">The flags (A to H, and I with its reserved bits when <paramref name="FlagsLength"/> is 2) as they
/// were sent.</param>
/// <param name="Versioning">The data of the Query Changes Versioning, as it was sent.</param>
record QueryChangesRequest(
    ulong RequestId,
    ushort Flags,
    int FlagsLength = 1,
    QueryChangesArguments? Arguments = null,
    ulong? MaxDataElements = null,
    ReadOnlyMemory<byte>? Versioning = null,
    IReadOnlyList<QueryChangesFilter>? Filters = null,
    Knowledge? Knowledge = null,
    ulong Priority = 0,
    Guid? TargetPartition = null) :
    SubRequest(RequestId, Priority, TargetPartition)
{
    public override RequestType Type => RequestType.QueryChanges;

    public bool AllowFragments => (Flags & 0x02) != 0 || (Flags & 0x10) != 0;

    public bool IncludeFilteredOutDataElementsInKnowledge => (Flags & 0x08) != 0;

    public bool RoundKnowledgeToWholeCellChanges => (Flags & 0x20) != 0;

    public bool ReturnFileHash => (Flags & 0x40) != 0;

    public bool UserContentEquivalentVersionOk => FlagsLength > 1 && (Flags & 0x100) != 0;

    public static QueryChangesRequest ReadData(ref FsshttpbReader reader, ulong requestId, ulong priority, Guid? partition)
    {
        var flagsFields = reader.ReadStart(StreamObjectType.QueryChangesRequest);
        var flagsLength = flagsFields.Remaining;
        ushort flags;
        if (flagsLength == 1)
        {
            flags = flagsFields.ReadByte();
        }
        else if (flagsLength == 2)
        {
            flags = flagsFields.ReadUInt16();
        }
        else
        {
            throw flagsFields.Error($"Query Changes flags of {flagsLength} bytes");
        }

        QueryChangesArguments? arguments = null;
        if (reader.PeekStart(StreamObjectType.QueryChangesRequestArguments))
        {
            var fields = reader.ReadStart(StreamObjectType.QueryChangesRequestArguments);
            var argumentFlags = fields.ReadByte();
            CellId? cell = null;
            if (!fields.AtEnd)
            {
                cell = fields.ReadCellId();
            }

            fields.EnsureEnd("the Query Changes request arguments");
            arguments = new(argumentFlags, cell);
        }

        ulong? maxDataElements = null;
        if (reader.PeekStart(StreamObjectType.QueryChangesDataConstraint))
        {
            var fields = reader.ReadStart(StreamObjectType.QueryChangesDataConstraint);
            maxDataElements = fields.ReadCompactUInt64();
            fields.EnsureEnd("the Query Changes data constraints");
        }

        ReadOnlyMemory<byte>? versioning = null;
        if (reader.PeekStart(StreamObjectType.QueryChangesVersioning))
        {
            versioning = reader.ReadStart(StreamObjectType.QueryChangesVersioning).ReadRest();
        }

        List<QueryChangesFilter>? filters = null;
        while (reader.PeekStart(StreamObjectType.QueryChangesFilter))
        {
            filters ??= [];
            filters.Add(ReadFilter(ref reader));
        }

        Knowledge? knowledge = null;
        if (reader.PeekStart(StreamObjectType.Knowledge))
        {
            knowledge = Knowledge.Read(ref reader);
        }

        return new(requestId, flags, flagsLength, arguments, maxDataElements, versioning, filters, knowledge, priority, partition);
    }

    static QueryChangesFilter ReadFilter(ref FsshttpbReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.QueryChangesFilter);
        var filterType = fields.ReadByte();
        var operation = fields.ReadByte();
        fields.EnsureEnd("a Query Changes filter start");
        RawStreamObject? data = null;
        if (!reader.PeekEnd())
        {
            data = reader.ReadSingle();
        }

        reader.ReadEnd(StreamObjectType.QueryChangesFilter);
        byte? flags = null;
        if (reader.PeekStart(StreamObjectType.QueryChangesFilterFlags))
        {
            var flagsFields = reader.ReadStart(StreamObjectType.QueryChangesFilterFlags);
            flags = flagsFields.ReadByte();
            flagsFields.EnsureEnd("Query Changes filter flags");
        }

        return new(filterType, operation, data, flags);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.QueryChangesRequest, FlagsLength);
        if (FlagsLength == 2)
        {
            writer.WriteUInt16(Flags);
        }
        else
        {
            writer.WriteByte((byte) Flags);
        }

        if (Arguments != null)
        {
            writer.WriteStart(StreamObjectType.QueryChangesRequestArguments, 1 + (Arguments.Cell?.Size ?? 0));
            writer.WriteByte(Arguments.Flags);
            if (Arguments.Cell != null)
            {
                writer.WriteCellId(Arguments.Cell.Value);
            }
        }

        if (MaxDataElements != null)
        {
            writer.WriteStart(StreamObjectType.QueryChangesDataConstraint, CompactUInt64.Size(MaxDataElements.Value));
            writer.WriteCompactUInt64(MaxDataElements.Value);
        }

        if (Versioning != null)
        {
            writer.WriteStart(StreamObjectType.QueryChangesVersioning, Versioning.Value.Length);
            writer.WriteBytes(Versioning.Value.Span);
        }

        foreach (var filter in Filters ?? [])
        {
            writer.WriteStart(StreamObjectType.QueryChangesFilter, 2);
            writer.WriteByte(filter.FilterType);
            writer.WriteByte(filter.Operation);
            if (filter.Data != null)
            {
                writer.WriteSingle(filter.Data);
            }

            writer.WriteEnd(StreamObjectType.QueryChangesFilter);
            if (filter.Flags != null)
            {
                writer.WriteStart(StreamObjectType.QueryChangesFilterFlags, 1);
                writer.WriteByte(filter.Flags.Value);
            }
        }

        Knowledge?.Write(ref writer);
    }
}

/// <summary>
/// A Put Changes sub-request ([MS-FSSHTTPB] 2.2.2.1.4): the Storage Index to apply, from the Data Element Package of the
/// request, and the one the client expects the server to have.
/// </summary>
/// <param name="Flags">A, imply null expected if no mapping (0x01); B, partial (0x02); C, partial last (0x04); D,
/// favor coherency failure over not found (0x08); E, abort remaining Put Changes on failure (0x10); F, multi-request
/// put hint (0x20); G, return complete knowledge if possible (0x40); H, last writer wins on next change (0x80).</param>
/// <param name="Extra">The optional fields of the Put Changes request object after the flags (content version
/// coherency check, author logins, reserved), as they were sent.</param>
/// <param name="AdditionalFlags">The data of the Additional Flags (2.2.2.1.4.1), as it was sent.</param>
/// <param name="DiagnosticInput">The flags of the Diagnostic Request Option Input (2.2.2.1.4.3).</param>
record PutChangesRequest(
    ulong RequestId,
    ExtendedGuid StorageIndex,
    ExtendedGuid ExpectedStorageIndex,
    byte Flags,
    ReadOnlyMemory<byte> Extra = default,
    ReadOnlyMemory<byte>? AdditionalFlags = null,
    Guid? LockId = null,
    Knowledge? ClientKnowledge = null,
    byte? DiagnosticInput = null,
    ulong Priority = 0,
    Guid? TargetPartition = null) :
    SubRequest(RequestId, Priority, TargetPartition)
{
    public override RequestType Type => RequestType.PutChanges;

    public bool ImplyNullExpectedIfNoMapping => (Flags & 0x01) != 0;

    public bool Partial => (Flags & 0x02) != 0;

    public bool PartialLast => (Flags & 0x04) != 0;

    public bool FavorCoherencyFailureOverNotFound => (Flags & 0x08) != 0;

    public bool LastWriterWinsOnNextChange => (Flags & 0x80) != 0;

    /// <summary>
    /// The flags of the Additional Flags (2.2.2.1.4.1): A, return applied Storage Index id entries (0x01); B, return data
    /// elements added (0x02); C, check for id reuse (0x04); D, coherency check only applied index entries (0x08); E,
    /// full file replace put (0x10); F, require storage mappings rooted (0x20).
    /// </summary>
    public ushort AdditionalFlagBits
    {
        get
        {
            if (AdditionalFlags is { Length: >= 2 } flags)
            {
                return BinaryPrimitives.ReadUInt16LittleEndian(flags.Span);
            }

            return 0;
        }
    }

    public static PutChangesRequest ReadData(ref FsshttpbReader reader, ulong requestId, ulong priority, Guid? partition)
    {
        var fields = reader.ReadStart(StreamObjectType.PutChangesRequest);
        var storageIndex = fields.ReadExtendedGuid();
        var expected = fields.ReadExtendedGuid();
        var flags = fields.ReadByte();
        var extra = fields.ReadRest();
        ReadOnlyMemory<byte>? additionalFlags = null;
        if (reader.PeekStart(StreamObjectType.AdditionalFlags))
        {
            additionalFlags = reader.ReadStart(StreamObjectType.AdditionalFlags).ReadRest();
        }

        Guid? lockId = null;
        if (reader.PeekStart(StreamObjectType.PutChangesLockId))
        {
            var lockFields = reader.ReadStart(StreamObjectType.PutChangesLockId);
            lockId = lockFields.ReadGuid();
            lockFields.EnsureEnd("a Put Changes lock id");
        }

        Knowledge? knowledge = null;
        if (reader.PeekStart(StreamObjectType.Knowledge))
        {
            knowledge = Knowledge.Read(ref reader);
        }

        byte? diagnostic = null;
        if (reader.PeekStart(StreamObjectType.DiagnosticRequestOptionInput))
        {
            var diagnosticFields = reader.ReadStart(StreamObjectType.DiagnosticRequestOptionInput);
            diagnostic = diagnosticFields.ReadByte();
            diagnosticFields.EnsureEnd("a diagnostic request option input");
        }

        return new(requestId, storageIndex, expected, flags, extra, additionalFlags, lockId, knowledge, diagnostic, priority, partition);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.PutChangesRequest, StorageIndex.Size + ExpectedStorageIndex.Size + 1 + Extra.Length);
        writer.WriteExtendedGuid(StorageIndex);
        writer.WriteExtendedGuid(ExpectedStorageIndex);
        writer.WriteByte(Flags);
        writer.WriteBytes(Extra.Span);
        if (AdditionalFlags != null)
        {
            writer.WriteStart(StreamObjectType.AdditionalFlags, AdditionalFlags.Value.Length);
            writer.WriteBytes(AdditionalFlags.Value.Span);
        }

        if (LockId != null)
        {
            writer.WriteStart(StreamObjectType.PutChangesLockId, 16);
            writer.WriteGuid(LockId.Value);
        }

        ClientKnowledge?.Write(ref writer);
        if (DiagnosticInput != null)
        {
            writer.WriteStart(StreamObjectType.DiagnosticRequestOptionInput, 1);
            writer.WriteByte(DiagnosticInput.Value);
        }
    }
}

/// <summary>
/// An Allocate Extended GUID Range sub-request ([MS-FSSHTTPB] 2.2.2.1.5): the number of Extended GUIDs to allocate.
/// </summary>
record AllocateExtendedGuidRangeRequest(ulong RequestId, ulong Count, ulong Priority = 0, Guid? TargetPartition = null) :
    SubRequest(RequestId, Priority, TargetPartition)
{
    public override RequestType Type => RequestType.AllocateExtendedGuidRange;

    public static AllocateExtendedGuidRangeRequest ReadData(ref FsshttpbReader reader, ulong requestId, ulong priority, Guid? partition)
    {
        var fields = reader.ReadStart(StreamObjectType.AllocateExtendedGuidRangeRequest);
        var count = fields.ReadCompactUInt64();
        // the reserved byte, which is ignored
        fields.ReadByte();
        fields.EnsureEnd("an Allocate Extended GUID Range request");
        return new(requestId, count, priority, partition);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.AllocateExtendedGuidRangeRequest, CompactUInt64.Size(Count) + 1);
        writer.WriteCompactUInt64(Count);
        writer.WriteByte(0);
    }
}
