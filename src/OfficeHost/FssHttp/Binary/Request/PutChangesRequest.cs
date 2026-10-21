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

    public bool CheckForIdReuse => (AdditionalFlagBits & 0x04) != 0;

    public static PutChangesRequest ReadData(ref FssHttpBReader reader, ulong requestId, ulong priority, Guid? partition)
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

    protected override void WriteData(ref FssHttpBWriter writer)
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