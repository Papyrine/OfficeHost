/// <summary>
/// The Storage Index Manifest Mapping: the Storage Manifest.
/// </summary>
record StorageIndexManifestMapping(ExtendedGuid MappingId, SerialNumber MappingSerialNumber) :
    StorageIndexEntry(MappingId, MappingSerialNumber);