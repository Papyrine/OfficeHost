/// <summary>
/// The File Hash of a Query Changes sub-response: the type of the hash and the hash.
/// </summary>
record FileHash(ulong Type, ReadOnlyMemory<byte> Data);