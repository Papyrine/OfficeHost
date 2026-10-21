/// <summary>
/// A Waterline Knowledge Entry ([MS-FSSHTTPB] 2.2.1.13.4.1): the waterline of a cell storage, opaque to clients.
/// </summary>
record WaterlineKnowledgeEntry(ExtendedGuid CellStorage, ulong Waterline, ulong Reserved = 0);