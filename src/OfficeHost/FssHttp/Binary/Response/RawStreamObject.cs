/// <summary>
/// A single stream object kept as its type and data, for the parts of the protocol the codec carries through without
/// interpreting.
/// </summary>
record RawStreamObject(StreamObjectType Type, ReadOnlyMemory<byte> Data);