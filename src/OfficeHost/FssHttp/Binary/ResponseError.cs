/// <summary>
/// The error types of [MS-FSSHTTPB] 2.2.3.2.
/// </summary>
enum ResponseErrorType
{
    Cell,
    Protocol,
    Win32,
    HResult
}

/// <summary>
/// The Cell Error codes of [MS-FSSHTTPB] 2.2.3.2.1 that a server uses.
/// </summary>
enum CellErrorCode : uint
{
    Unknown = 1,
    InvalidObject = 2,
    InvalidPartition = 3,
    RequestNotSupported = 4,
    StorageReadOnly = 5,
    RevisionIdNotFound = 6,
    BadToken = 7,
    RequestNotFinished = 8,
    IncompatibleToken = 9,
    ScopedCellStorage = 11,
    CoherencyFailure = 12,
    CellStorageStateDeserializationFailure = 13,
    IncompatibleProtocolVersion = 15,
    ReferencedDataElementNotFound = 16,
    RequestStreamSchemaError = 18,
    ResponseStreamSchemaError = 19,
    UnknownRequest = 20,
    StorageFailure = 21,
    StorageWriteOnly = 22,
    InvalidSerialization = 23,
    DataElementNotFound = 24,
    InvalidImplementation = 25,
    IncompatibleOldStorage = 26,
    IncompatibleNewStorage = 27,
    IncorrectContextForDataElementId = 28,
    ObjectGroupDuplicateObjects = 29,
    ObjectReferenceNotFoundInRevision = 31,
    MergeCellStorageStateConflict = 32,
    UnknownQueryChangesFilter = 33,
    UnsupportedQueryChangesFilter = 34,
    UnableToProvideKnowledge = 35,
    DataElementMissingId = 36,
    DataElementMissingSerialNumber = 37,
    RequestArgumentInvalid = 38,
    PartialChangesNotSupported = 39,
    StoreBusyRetryLater = 40,
    GuidIdTableNotSupported = 41,
    DataElementCycle = 42,
    FragmentKnowledgeError = 43,
    FragmentSizeMismatch = 44,
    FragmentsIncomplete = 45,
    FragmentInvalid = 46,
    AbortedAfterFailedPutChanges = 47
}

/// <summary>
/// A Response Error ([MS-FSSHTTPB] 2.2.3.2): the type of the error and its code, an optional supplemental string, and an
/// optional chained error.
/// </summary>
record ResponseError(ResponseErrorType Type, uint Code, string? SupplementalInfo = null, ResponseError? Chained = null)
{
    const int maxChain = 8;

    /// <summary>
    /// The HRESULT error zero: the success of a Query Access check (2.2.3.1.1).
    /// </summary>
    public static ResponseError Success { get; } = new(ResponseErrorType.HResult, 0);

    public static ResponseError Cell(CellErrorCode code) => new(ResponseErrorType.Cell, (uint) code);

    static Guid cellGuid = new("5A66A756-87CE-4290-A38B-C61C5BA05A67");
    static Guid protocolGuid = new("7AFEAEBF-033D-4828-9C31-3977AFE58249");
    static Guid win32Guid = new("32C39011-6E39-46C4-AB78-DB41929D679E");
    static Guid hresultGuid = new("8454C8F2-E401-405A-A198-A10B6991B56E");

    static (Guid Guid, StreamObjectType Header) Describe(ResponseErrorType type) =>
        type switch
        {
            ResponseErrorType.Cell => (cellGuid, StreamObjectType.ErrorCell),
            ResponseErrorType.Protocol => (protocolGuid, StreamObjectType.ErrorProtocol),
            ResponseErrorType.Win32 => (win32Guid, StreamObjectType.ErrorWin32),
            _ => (hresultGuid, StreamObjectType.ErrorHresult)
        };

    public static ResponseError Read(ref FssHttpBReader reader) => Read(ref reader, 0);

    static ResponseError Read(ref FssHttpBReader reader, int depth)
    {
        var at = reader.Offset;
        if (depth > maxChain)
        {
            throw reader.Error("Response errors chained too deeply");
        }

        var fields = reader.ReadStart(StreamObjectType.Error);
        var guid = fields.ReadGuid();
        fields.EnsureEnd("an error start");
        ResponseErrorType type;
        if (guid == cellGuid)
        {
            type = ResponseErrorType.Cell;
        }
        else if (guid == protocolGuid)
        {
            type = ResponseErrorType.Protocol;
        }
        else if (guid == win32Guid)
        {
            type = ResponseErrorType.Win32;
        }
        else if (guid == hresultGuid)
        {
            type = ResponseErrorType.HResult;
        }
        else
        {
            throw new FssHttpBException($"Unknown error type {guid}", at);
        }

        var data = reader.ReadStart(Describe(type).Header);
        var code = data.ReadUInt32();
        data.EnsureEnd("an error code");
        string? supplemental = null;
        if (reader.PeekStart(StreamObjectType.ErrorStringSupplementalInfo))
        {
            var info = reader.ReadStart(StreamObjectType.ErrorStringSupplementalInfo);
            supplemental = info.ReadStringItem();
            info.EnsureEnd("an error string supplemental info");
        }

        ResponseError? chained = null;
        if (reader.PeekStart(StreamObjectType.Error))
        {
            chained = Read(ref reader, depth + 1);
        }

        reader.ReadEnd(StreamObjectType.Error);
        return new(type, code, supplemental, chained);
    }

    public void Write(ref FssHttpBWriter writer)
    {
        var (guid, header) = Describe(Type);
        writer.WriteStart(StreamObjectType.Error, 16);
        writer.WriteGuid(guid);
        writer.WriteStart(header, 4);
        writer.WriteUInt32(Code);
        if (SupplementalInfo != null)
        {
            writer.WriteStart(StreamObjectType.ErrorStringSupplementalInfo, FssHttpBSize.StringItem(SupplementalInfo));
            writer.WriteStringItem(SupplementalInfo);
        }

        Chained?.Write(ref writer);
        writer.WriteEnd(StreamObjectType.Error);
    }
}
