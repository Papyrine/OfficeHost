/// <summary>
/// The state of the cell storage of a partition of a file, as an <see cref="IFssHttpStore"/> keeps it: the data elements,
/// the next identifier to allocate, so that Serial Numbers keep increasing, and for the content partition the version of
/// the file in the <see cref="IOfficeStorage"/> the cell holds.
/// </summary>
/// <param name="Storage">The data elements, null for a partition Office has not written yet.</param>
/// <param name="FileVersion">The <see cref="OfficeFile.Version"/> the cell holds, null for the other partitions.</param>
/// <param name="BaseFileVersion">For a content cell a Put Changes stored before writing its file (with no
/// <paramref name="FileVersion"/>), the version of the file the cell was changed from: once the file has another, the
/// cell is out of date and not written.</param>
record CellState(CellStorage? Storage, Guid IdGuid, ulong NextId, string? FileVersion, string? BaseFileVersion = null)
{
    // "OHC" and the version of the format: 2 adds the base file version
    static ReadOnlySpan<byte> magic => "OHC2"u8;

    static ReadOnlySpan<byte> magic1 => "OHC1"u8;

    public static CellState New() => new(null, Guid.NewGuid(), 1, null);

    public IdAllocator Ids() => new(IdGuid, NextId);

    /// <summary>
    /// The magic and format version, the file version and the base file version as optional String Items, the GUID and
    /// next value of the identifiers, then the Storage Index and a Data Element Package ([MS-FSSHTTPB] 2.2.1), when there
    /// is a storage.
    /// </summary>
    public byte[] ToArray()
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        writer.WriteBytes(magic);
        WriteOptional(ref writer, FileVersion != null);
        if (FileVersion != null)
        {
            writer.WriteStringItem(FileVersion);
        }

        WriteOptional(ref writer, BaseFileVersion != null);
        if (BaseFileVersion != null)
        {
            writer.WriteStringItem(BaseFileVersion);
        }

        writer.WriteGuid(IdGuid);
        writer.WriteCompactUInt64(NextId);
        WriteOptional(ref writer, Storage != null);
        if (Storage != null)
        {
            writer.WriteExtendedGuid(Storage.StorageIndexId);
            new DataElementPackage(Storage.Elements.Values.ToList()).Write(ref writer);
        }

        return output.WrittenSpan.ToArray();
    }

    static void WriteOptional(ref FssHttpBWriter writer, bool present)
    {
        byte value = 0;
        if (present)
        {
            value = 1;
        }

        writer.WriteByte(value);
    }

    /// <summary>
    /// The state of <see cref="ToArray"/>, whose data elements are slices of <paramref name="data"/>.
    /// </summary>
    public static CellState Read(ReadOnlyMemory<byte> data)
    {
        try
        {
            var reader = new FssHttpBReader(data);
            var format = reader.ReadBytes(4).Span;
            var withBase = format.SequenceEqual(magic);
            if (!withBase &&
                !format.SequenceEqual(magic1))
            {
                throw new InvalidDataException("Not the state of a cell");
            }

            string? fileVersion = null;
            if (ReadOptional(ref reader))
            {
                fileVersion = reader.ReadStringItem();
            }

            string? baseFileVersion = null;
            if (withBase &&
                ReadOptional(ref reader))
            {
                baseFileVersion = reader.ReadStringItem();
            }

            var idGuid = reader.ReadGuid();
            var nextId = reader.ReadCompactUInt64();
            CellStorage? storage = null;
            if (ReadOptional(ref reader))
            {
                var storageIndexId = reader.ReadExtendedGuid();
                storage = CellStorage.Create(DataElementPackage.Read(ref reader).Elements, storageIndexId);
            }

            reader.EnsureEnd("the state of a cell");
            return new(storage, idGuid, nextId, fileVersion, baseFileVersion);
        }
        catch (FssHttpBException exception)
        {
            throw new InvalidDataException($"The state of a cell is corrupt: {exception.Message}", exception);
        }
    }

    static bool ReadOptional(ref FssHttpBReader reader) =>
        reader.ReadByte() switch
        {
            0 => false,
            1 => true,
            var value => throw reader.Error($"Invalid flag {value}")
        };
}
