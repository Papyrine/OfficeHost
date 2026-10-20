/// <summary>
/// FSSHTTPB ([MS-FSSHTTPB]) input that is malformed, truncated, or uses a structure this codec does not know: the only
/// exception the readers throw for bad input.
/// </summary>
class FssHttpBException(string message, int offset) :
    Exception($"{message} (at byte {offset})")
{
    /// <summary>
    /// The position in the input where the problem was found.
    /// </summary>
    public int Offset { get; } = offset;
}
