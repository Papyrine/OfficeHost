/// <summary>
/// The Request Hashing Options of a request ([MS-FSSHTTPB] 2.2.2): the hashing schema, and a byte of flags (C, request
/// data element hashes instead of data, 0x04; D, request data element hashes, 0x08).
/// </summary>
record RequestHashingOptions(ulong Schema, byte Flags);