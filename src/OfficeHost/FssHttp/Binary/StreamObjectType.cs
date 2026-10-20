/// <summary>
/// The stream object types of [MS-FSSHTTPB] 2.2.1.5.1 (below 0x40, for 16-bit headers) and 2.2.1.5.2 (32-bit headers),
/// and those of the node objects of [MS-FSSHTTPD] 2.2.
/// </summary>
enum StreamObjectType : ushort
{
    DataElement = 0x01,
    ObjectDataBlob = 0x02,
    ObjectGroupObjectExcludedData = 0x03,
    WaterlineKnowledgeEntry = 0x04,
    ObjectGroupObjectDataBlobDeclaration = 0x05,
    DataElementHash = 0x06,
    StorageManifestRootDeclare = 0x07,
    RevisionManifestRootDeclare = 0x0A,
    CellManifestCurrentRevision = 0x0B,
    StorageManifestSchemaGuid = 0x0C,
    StorageIndexRevisionMapping = 0x0D,
    StorageIndexCellMapping = 0x0E,
    CellKnowledgeRange = 0x0F,
    Knowledge = 0x10,
    StorageIndexManifestMapping = 0x11,
    CellKnowledge = 0x14,
    DataElementPackage = 0x15,
    ObjectGroupObjectData = 0x16,
    CellKnowledgeEntry = 0x17,
    ObjectGroupObjectDeclare = 0x18,
    RevisionManifestObjectGroupReferences = 0x19,
    RevisionManifest = 0x1A,
    ObjectGroupObjectDataBlobReference = 0x1C,
    ObjectGroupDeclarations = 0x1D,
    ObjectGroupData = 0x1E,
    // [MS-FSSHTTPD] 2.2.2.1 and 2.2.3.1
    LeafNode = 0x1F,
    IntermediateNode = 0x20,
    Signature = 0x21,
    DataSize = 0x22,
    WaterlineKnowledge = 0x29,
    // [MS-FSSHTTPD] 2.2.3.1
    DataHash = 0x2F,
    ContentTagKnowledge = 0x2D,
    ContentTagKnowledgeEntry = 0x2E,
    QueryChangesVersioning = 0x30,
    Request = 0x40,
    SubResponse = 0x41,
    SubRequest = 0x42,
    ReadAccessResponse = 0x43,
    SpecializedKnowledge = 0x44,
    WriteAccessResponse = 0x46,
    QueryChangesFilter = 0x47,
    ErrorWin32 = 0x49,
    ErrorProtocol = 0x4B,
    Error = 0x4D,
    ErrorStringSupplementalInfo = 0x4E,
    UserAgentVersion = 0x4F,
    QueryChangesFilterSchemaSpecific = 0x50,
    QueryChangesRequest = 0x51,
    ErrorHresult = 0x52,
    QueryChangesFilterDataElementIds = 0x54,
    UserAgentGuid = 0x55,
    QueryChangesFilterDataElementType = 0x57,
    QueryChangesDataConstraint = 0x59,
    PutChangesRequest = 0x5A,
    QueryChangesRequestArguments = 0x5B,
    QueryChangesFilterCellId = 0x5C,
    UserAgent = 0x5D,
    QueryChangesResponse = 0x5F,
    QueryChangesFilterHierarchy = 0x60,
    Response = 0x62,
    ErrorCell = 0x66,
    QueryChangesFilterFlags = 0x68,
    DataElementFragment = 0x6A,
    FragmentKnowledge = 0x6B,
    FragmentKnowledgeEntry = 0x6C,
    ObjectGroupMetadata = 0x78,
    ObjectGroupMetadataDeclarations = 0x79,
    AllocateExtendedGuidRangeRequest = 0x80,
    AllocateExtendedGuidRangeResponse = 0x81,
    TargetPartitionId = 0x83,
    PutChangesLockId = 0x85,
    AdditionalFlags = 0x86,
    PutChangesResponse = 0x87,
    RequestHashingOptions = 0x88,
    DiagnosticRequestOptionOutput = 0x89,
    DiagnosticRequestOptionInput = 0x8A,
    UserAgentClientAndPlatform = 0x8B,
    VersionTokenKnowledge = 0x8C,
    CellRoundtripOptions = 0x8D,
    FileHash = 0x8E
}

static class StreamObjectTypes
{
    /// <summary>
    /// Whether objects of the type are compound, with an end header ([MS-FSSHTTPB] 2.2.1.5.1 and 2.2.1.5.2, the Compound
    /// column).
    /// </summary>
    public static bool IsCompound(this StreamObjectType type) =>
        type is StreamObjectType.DataElement or
            StreamObjectType.Knowledge or
            StreamObjectType.CellKnowledge or
            StreamObjectType.DataElementPackage or
            StreamObjectType.ObjectGroupDeclarations or
            StreamObjectType.ObjectGroupData or
            StreamObjectType.LeafNode or
            StreamObjectType.IntermediateNode or
            StreamObjectType.WaterlineKnowledge or
            StreamObjectType.ContentTagKnowledge or
            StreamObjectType.Request or
            StreamObjectType.SubResponse or
            StreamObjectType.SubRequest or
            StreamObjectType.ReadAccessResponse or
            StreamObjectType.SpecializedKnowledge or
            StreamObjectType.WriteAccessResponse or
            StreamObjectType.QueryChangesFilter or
            StreamObjectType.Error or
            StreamObjectType.UserAgent or
            StreamObjectType.Response or
            StreamObjectType.FragmentKnowledge or
            StreamObjectType.ObjectGroupMetadataDeclarations;
}
