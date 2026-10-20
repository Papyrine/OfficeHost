/// <summary>
/// Allocates the identifiers of new objects and data elements, and the Serial Numbers of new data elements and mappings
/// ([MS-FSSHTTPB] 2.2.1.7, 2.2.1.9): one GUID with increasing values, so that the Serial Numbers of a cell storage make
/// ranges for Cell Knowledge (2.2.1.13.2.1). To keep them increasing, a store keeps the GUID and the next value.
/// </summary>
class IdAllocator(Guid guid, ulong next = 1)
{
    public Guid Guid => guid;

    public ulong Next => next;

    public ExtendedGuid NewId()
    {
        if (next > uint.MaxValue)
        {
            throw new InvalidOperationException($"The values of the Extended GUIDs of {guid} are used up");
        }

        return new(guid, (uint) next++);
    }

    public SerialNumber NewSerialNumber() => new(guid, next++);
}
