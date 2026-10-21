/// <summary>
/// A Storage Index Revision Mapping: the Revision Manifest of a revision.
/// </summary>
record StorageIndexRevisionMapping(ExtendedGuid Revision, ExtendedGuid MappingId, SerialNumber MappingSerialNumber) :
    StorageIndexEntry(MappingId, MappingSerialNumber);