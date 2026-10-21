/// <summary>
/// A Data Element Hash ([MS-FSSHTTPB] 2.2.1.12.6.6): the hash scheme (1, [MS-PCCRC] 2.3) and the hash.
/// </summary>
record DataElementHash(ulong Scheme, ReadOnlyMemory<byte> Data);