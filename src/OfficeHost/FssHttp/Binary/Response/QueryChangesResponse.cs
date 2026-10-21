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