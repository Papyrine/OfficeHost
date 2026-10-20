/// <summary>
/// A stream object header ([MS-FSSHTTPB] 2.2.1.5): the 16-bit (2.2.1.5.1) or 32-bit (2.2.1.5.2) start of a single or
/// compound stream object, with the length of the data that follows it before the next header, or the 8-bit
/// (2.2.1.5.3) or 16-bit (2.2.1.5.4) end of a compound one.
/// </summary>
readonly record struct StreamObjectHeader(StreamObjectType Type, bool Compound, bool End, ulong Length)
{
    public override string ToString()
    {
        if (End)
        {
            return $"the end of {Type}";
        }

        return $"the start of {Type} ({Length} bytes)";
    }
}
