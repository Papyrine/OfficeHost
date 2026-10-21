/// <summary>
/// A mapping of a Storage Index ([MS-FSSHTTPB] 2.2.1.12.2), each with the Serial Number of the mapping.
/// </summary>
abstract record StorageIndexEntry(ExtendedGuid MappingId, SerialNumber MappingSerialNumber);