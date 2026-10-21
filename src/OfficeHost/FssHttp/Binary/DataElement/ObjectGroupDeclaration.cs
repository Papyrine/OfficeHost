/// <summary>
/// A declaration of an Object Group: an Object Declaration (2.2.1.12.6.1) or an Object Data BLOB Declaration
/// (2.2.1.12.6.2).
/// </summary>
abstract record ObjectGroupDeclaration(ExtendedGuid ObjectId, ulong PartitionId, ulong ObjectReferencesCount, ulong CellReferencesCount);