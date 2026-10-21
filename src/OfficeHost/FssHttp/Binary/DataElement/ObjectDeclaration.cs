/// <summary>
/// An Object Declaration ([MS-FSSHTTPB] 2.2.1.12.6.1).
/// </summary>
record ObjectDeclaration(ExtendedGuid ObjectId, ulong PartitionId, ulong DataSize, ulong ObjectReferencesCount, ulong CellReferencesCount) :
    ObjectGroupDeclaration(ObjectId, PartitionId, ObjectReferencesCount, CellReferencesCount);