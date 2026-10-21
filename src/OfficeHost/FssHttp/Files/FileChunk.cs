/// <summary>
/// A chunk of a file ([MS-FSSHTTPD] 2.4): a range of the file and its signature, and the subchunks of a chunk larger than
/// 3 megabytes (2.4.1).
/// </summary>
record FileChunk(int Start, int Length, ReadOnlyMemory<byte> Signature, IReadOnlyList<FileChunk> SubChunks);