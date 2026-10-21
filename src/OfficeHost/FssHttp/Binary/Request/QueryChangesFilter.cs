/// <summary>
/// A filter of a Query Changes sub-request ([MS-FSSHTTPB] 2.2.2.1.3.1): its type (1 all, 2 data element type, 3
/// Storage Index referenced, 4 Cell ID, 5 custom, 6 data element IDs, 7 hierarchy), whether it includes (1) or
/// excludes (0), its data as the stream object of the type, and its optional flags.
/// </summary>
record QueryChangesFilter(byte FilterType, byte Operation, RawStreamObject? Data, byte? Flags);