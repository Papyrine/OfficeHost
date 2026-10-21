/// <summary>
/// An Object Data BLOB Declaration ([MS-FSSHTTPB] 2.2.1.12.6.2): an object whose data is in an Object Data BLOB.
/// </summary>
record ObjectDataBlobDeclaration(ExtendedGuid ObjectId, ExtendedGuid BlobId, ulong PartitionId, ulong ObjectReferencesCount, ulong CellReferencesCount) :
    ObjectGroupDeclaration(ObjectId, PartitionId, ObjectReferencesCount, CellReferencesCount);