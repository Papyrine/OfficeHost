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

    public static QueryChangesRequest ReadData(ref FssHttpBReader reader, ulong requestId, ulong priority, Guid? partition)
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

    static QueryChangesFilter ReadFilter(ref FssHttpBReader reader)
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

    protected override void WriteData(ref FssHttpBWriter writer)
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