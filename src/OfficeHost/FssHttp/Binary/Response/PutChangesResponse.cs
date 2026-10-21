/// <summary>
/// A Put Changes sub-response ([MS-FSSHTTPB] 2.2.3.1.3): the knowledge of the server after the changes. The Put Changes
/// response object, with the applied Storage Index and the data elements added, is read as optional: the example of
/// [MS-FSSHTTPB] 4.4 has none (docs/fsshttp-notes.md).
/// </summary>
/// <param name="Extra">What the Put Changes response object has after the data elements added, which the spec does not
/// define: SharePoint Server SE sends 29 bytes there (docs/fsshttp-notes.md). Kept as it was sent.</param>
record PutChangesResponse(
    ulong RequestId,
    Knowledge Knowledge,
    bool HasResponseObject = false,
    ExtendedGuid? AppliedStorageIndex = null,
    IReadOnlyList<ExtendedGuid>? DataElementsAdded = null,
    byte? DiagnosticOutput = null,
    ReadOnlyMemory<byte> Extra = default) :
    SubResponse(RequestId)
{
    public override RequestType Type => RequestType.PutChanges;

    public static PutChangesResponse ReadData(ref FssHttpBReader reader, ulong requestId)
    {
        var hasObject = false;
        ExtendedGuid? applied = null;
        ExtendedGuid[]? added = null;
        ReadOnlyMemory<byte> extra = default;
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

            extra = fields.ReadRest();
        }

        var knowledge = Knowledge.Read(ref reader);
        byte? diagnostic = null;
        if (reader.PeekStart(StreamObjectType.DiagnosticRequestOptionOutput))
        {
            var fields = reader.ReadStart(StreamObjectType.DiagnosticRequestOptionOutput);
            diagnostic = fields.ReadByte();
            fields.EnsureEnd("a diagnostic request option output");
        }

        return new(requestId, knowledge, hasObject, applied, added, diagnostic, extra);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        var hasExtra = !Extra.IsEmpty;
        if (HasResponseObject || AppliedStorageIndex != null || DataElementsAdded != null || hasExtra)
        {
            var size = 0;
            if (AppliedStorageIndex != null || DataElementsAdded != null || hasExtra)
            {
                size += (AppliedStorageIndex ?? ExtendedGuid.Null).Size;
            }

            if (DataElementsAdded != null || hasExtra)
            {
                size += FssHttpBSize.ExtendedGuidArray(DataElementsAdded ?? []);
            }

            writer.WriteStart(StreamObjectType.PutChangesResponse, size + Extra.Length);
            if (AppliedStorageIndex != null || DataElementsAdded != null || hasExtra)
            {
                writer.WriteExtendedGuid(AppliedStorageIndex ?? ExtendedGuid.Null);
            }

            if (DataElementsAdded != null || hasExtra)
            {
                writer.WriteExtendedGuidArray(DataElementsAdded ?? []);
            }

            writer.WriteBytes(Extra.Span);
        }

        Knowledge.Write(ref writer);
        if (DiagnosticOutput != null)
        {
            writer.WriteStart(StreamObjectType.DiagnosticRequestOptionOutput, 1);
            writer.WriteByte(DiagnosticOutput.Value);
        }
    }
}