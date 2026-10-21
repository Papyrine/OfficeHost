/// <summary>
/// A Content Tag Knowledge Entry ([MS-FSSHTTPB] 2.2.1.13.5.1): the clock data of a BLOB, opaque to clients.
/// </summary>
record ContentTagKnowledgeEntry(ExtendedGuid Blob, ReadOnlyMemory<byte> ClockData);