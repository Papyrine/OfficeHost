/// <summary>
/// A Fragment Knowledge Entry ([MS-FSSHTTPB] 2.2.1.13.3.1): a part of a large data element that is on the server.
/// </summary>
record FragmentKnowledgeEntry(ExtendedGuid DataElement, ulong DataElementSize, FileChunkReference Chunk);