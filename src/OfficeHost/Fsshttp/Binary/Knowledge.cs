/// <summary>
/// Knowledge ([MS-FSSHTTPB] 2.2.1.13): what a client knows of the state of a file, as Specialized Knowledge of several
/// kinds.
/// </summary>
record Knowledge(IReadOnlyList<SpecializedKnowledge> Items)
{
    public static Knowledge Empty { get; } = new([]);

    public static Knowledge Read(ref FsshttpbReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.Knowledge);
        fields.EnsureEnd("a Knowledge start");
        var items = new List<SpecializedKnowledge>();
        while (reader.PeekStart(StreamObjectType.SpecializedKnowledge))
        {
            items.Add(SpecializedKnowledge.Read(ref reader));
        }

        reader.ReadEnd(StreamObjectType.Knowledge);
        return new(items);
    }

    public void Write(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.Knowledge, 0);
        foreach (var item in Items)
        {
            item.Write(ref writer);
        }

        writer.WriteEnd(StreamObjectType.Knowledge);
    }
}

/// <summary>
/// A Specialized Knowledge ([MS-FSSHTTPB] 2.2.1.13.1): its kind, as a GUID, then its data.
/// </summary>
abstract record SpecializedKnowledge
{
    public static Guid CellKnowledgeKind { get; } = new("327A35F6-0761-4414-9686-51E900667A4D");
    public static Guid WaterlineKnowledgeKind { get; } = new("3A76E90E-8032-4D0C-B9DD-F3C65029433E");
    public static Guid FragmentKnowledgeKind { get; } = new("0ABE4F35-01DF-4134-A24A-7C79F0859844");
    public static Guid ContentTagKnowledgeKind { get; } = new("10091F13-C882-40FB-9886-6533F934C21D");
    public static Guid VersionTokenKnowledgeKind { get; } = new("BF12E2C1-E64F-4959-8282-73B9A24A7C44");

    public abstract Guid Kind { get; }

    public static SpecializedKnowledge Read(ref FsshttpbReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.SpecializedKnowledge);
        var kind = fields.ReadGuid();
        fields.EnsureEnd("a Specialized Knowledge start");
        SpecializedKnowledge knowledge;
        if (kind == CellKnowledgeKind)
        {
            knowledge = CellKnowledge.ReadData(ref reader);
        }
        else if (kind == WaterlineKnowledgeKind)
        {
            knowledge = WaterlineKnowledge.ReadData(ref reader);
        }
        else if (kind == FragmentKnowledgeKind)
        {
            knowledge = FragmentKnowledge.ReadData(ref reader);
        }
        else if (kind == ContentTagKnowledgeKind)
        {
            knowledge = ContentTagKnowledge.ReadData(ref reader);
        }
        else if (kind == VersionTokenKnowledgeKind)
        {
            knowledge = VersionTokenKnowledge.ReadData(ref reader);
        }
        else
        {
            knowledge = new UnknownKnowledge(kind, reader.ReadObjectsUntilEnd());
        }

        reader.ReadEnd(StreamObjectType.SpecializedKnowledge);
        return knowledge;
    }

    public void Write(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.SpecializedKnowledge, 16);
        writer.WriteGuid(Kind);
        WriteData(ref writer);
        writer.WriteEnd(StreamObjectType.SpecializedKnowledge);
    }

    protected abstract void WriteData(ref FsshttpbWriter writer);
}

/// <summary>
/// An item of Cell Knowledge: a Cell Knowledge Range (2.2.1.13.2.1) or a Cell Knowledge Entry (2.2.1.13.2.2).
/// </summary>
abstract record CellKnowledgeItem;

/// <summary>
/// A range of Serial Numbers: the GUID with the values from From to To.
/// </summary>
record CellKnowledgeRange(Guid Guid, ulong From, ulong To) :
    CellKnowledgeItem;

/// <summary>
/// The Serial Number of a single cell.
/// </summary>
record CellKnowledgeEntry(SerialNumber SerialNumber) :
    CellKnowledgeItem;

/// <summary>
/// Cell Knowledge ([MS-FSSHTTPB] 2.2.1.13.2): the data elements a client has, as ranges and single Serial Numbers.
/// </summary>
record CellKnowledge(IReadOnlyList<CellKnowledgeItem> Items) :
    SpecializedKnowledge
{
    public override Guid Kind => CellKnowledgeKind;

    public static CellKnowledge ReadData(ref FsshttpbReader reader)
    {
        reader.ReadStart(StreamObjectType.CellKnowledge).EnsureEnd("a Cell Knowledge start");
        var items = new List<CellKnowledgeItem>();
        while (!reader.PeekEnd())
        {
            var header = reader.PeekHeader();
            var fields = reader.ReadStart(header.Type);
            CellKnowledgeItem item = header.Type switch
            {
                StreamObjectType.CellKnowledgeRange => new CellKnowledgeRange(fields.ReadGuid(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64()),
                StreamObjectType.CellKnowledgeEntry => new CellKnowledgeEntry(fields.ReadSerialNumber()),
                _ => throw fields.Error($"{header} in Cell Knowledge")
            };
            fields.EnsureEnd("a Cell Knowledge item");
            items.Add(item);
        }

        reader.ReadEnd(StreamObjectType.CellKnowledge);
        return new(items);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.CellKnowledge, 0);
        foreach (var item in Items)
        {
            switch (item)
            {
                case CellKnowledgeRange range:
                    writer.WriteStart(StreamObjectType.CellKnowledgeRange, 16 + CompactUInt64.Size(range.From) + CompactUInt64.Size(range.To));
                    writer.WriteGuid(range.Guid);
                    writer.WriteCompactUInt64(range.From);
                    writer.WriteCompactUInt64(range.To);
                    break;
                case CellKnowledgeEntry entry:
                    writer.WriteStart(StreamObjectType.CellKnowledgeEntry, entry.SerialNumber.Size);
                    writer.WriteSerialNumber(entry.SerialNumber);
                    break;
            }
        }

        writer.WriteEnd(StreamObjectType.CellKnowledge);
    }
}

/// <summary>
/// A Waterline Knowledge Entry ([MS-FSSHTTPB] 2.2.1.13.4.1): the waterline of a cell storage, opaque to clients.
/// </summary>
record WaterlineKnowledgeEntry(ExtendedGuid CellStorage, ulong Waterline, ulong Reserved = 0);

/// <summary>
/// Waterline Knowledge ([MS-FSSHTTPB] 2.2.1.13.4).
/// </summary>
record WaterlineKnowledge(IReadOnlyList<WaterlineKnowledgeEntry> Entries) :
    SpecializedKnowledge
{
    public override Guid Kind => WaterlineKnowledgeKind;

    public static WaterlineKnowledge ReadData(ref FsshttpbReader reader)
    {
        reader.ReadStart(StreamObjectType.WaterlineKnowledge).EnsureEnd("a Waterline Knowledge start");
        var entries = new List<WaterlineKnowledgeEntry>();
        while (reader.PeekStart(StreamObjectType.WaterlineKnowledgeEntry))
        {
            var fields = reader.ReadStart(StreamObjectType.WaterlineKnowledgeEntry);
            entries.Add(new(fields.ReadExtendedGuid(), fields.ReadCompactUInt64(), fields.ReadCompactUInt64()));
            fields.EnsureEnd("a Waterline Knowledge Entry");
        }

        reader.ReadEnd(StreamObjectType.WaterlineKnowledge);
        return new(entries);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.WaterlineKnowledge, 0);
        foreach (var entry in Entries)
        {
            writer.WriteStart(StreamObjectType.WaterlineKnowledgeEntry, entry.CellStorage.Size + CompactUInt64.Size(entry.Waterline) + CompactUInt64.Size(entry.Reserved));
            writer.WriteExtendedGuid(entry.CellStorage);
            writer.WriteCompactUInt64(entry.Waterline);
            writer.WriteCompactUInt64(entry.Reserved);
        }

        writer.WriteEnd(StreamObjectType.WaterlineKnowledge);
    }
}

/// <summary>
/// A Fragment Knowledge Entry ([MS-FSSHTTPB] 2.2.1.13.3.1): a part of a large data element that is on the server.
/// </summary>
record FragmentKnowledgeEntry(ExtendedGuid DataElement, ulong DataElementSize, FileChunkReference Chunk);

/// <summary>
/// Fragment Knowledge ([MS-FSSHTTPB] 2.2.1.13.3).
/// </summary>
record FragmentKnowledge(IReadOnlyList<FragmentKnowledgeEntry> Entries) :
    SpecializedKnowledge
{
    public override Guid Kind => FragmentKnowledgeKind;

    public static FragmentKnowledge ReadData(ref FsshttpbReader reader)
    {
        reader.ReadStart(StreamObjectType.FragmentKnowledge).EnsureEnd("a Fragment Knowledge start");
        var entries = new List<FragmentKnowledgeEntry>();
        while (reader.PeekStart(StreamObjectType.FragmentKnowledgeEntry))
        {
            var fields = reader.ReadStart(StreamObjectType.FragmentKnowledgeEntry);
            entries.Add(new(fields.ReadExtendedGuid(), fields.ReadCompactUInt64(), fields.ReadFileChunkReference()));
            fields.EnsureEnd("a Fragment Knowledge Entry");
        }

        reader.ReadEnd(StreamObjectType.FragmentKnowledge);
        return new(entries);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.FragmentKnowledge, 0);
        foreach (var entry in Entries)
        {
            writer.WriteStart(StreamObjectType.FragmentKnowledgeEntry, entry.DataElement.Size + CompactUInt64.Size(entry.DataElementSize) + entry.Chunk.Size);
            writer.WriteExtendedGuid(entry.DataElement);
            writer.WriteCompactUInt64(entry.DataElementSize);
            writer.WriteFileChunkReference(entry.Chunk);
        }

        writer.WriteEnd(StreamObjectType.FragmentKnowledge);
    }
}

/// <summary>
/// A Content Tag Knowledge Entry ([MS-FSSHTTPB] 2.2.1.13.5.1): the clock data of a BLOB, opaque to clients.
/// </summary>
record ContentTagKnowledgeEntry(ExtendedGuid Blob, ReadOnlyMemory<byte> ClockData);

/// <summary>
/// Content Tag Knowledge ([MS-FSSHTTPB] 2.2.1.13.5).
/// </summary>
record ContentTagKnowledge(IReadOnlyList<ContentTagKnowledgeEntry> Entries) :
    SpecializedKnowledge
{
    public override Guid Kind => ContentTagKnowledgeKind;

    public static ContentTagKnowledge ReadData(ref FsshttpbReader reader)
    {
        reader.ReadStart(StreamObjectType.ContentTagKnowledge).EnsureEnd("a Content Tag Knowledge start");
        var entries = new List<ContentTagKnowledgeEntry>();
        while (reader.PeekStart(StreamObjectType.ContentTagKnowledgeEntry))
        {
            var fields = reader.ReadStart(StreamObjectType.ContentTagKnowledgeEntry);
            entries.Add(new(fields.ReadExtendedGuid(), fields.ReadBinaryItem()));
            fields.EnsureEnd("a Content Tag Knowledge Entry");
        }

        reader.ReadEnd(StreamObjectType.ContentTagKnowledge);
        return new(entries);
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.ContentTagKnowledge, 0);
        foreach (var entry in Entries)
        {
            writer.WriteStart(StreamObjectType.ContentTagKnowledgeEntry, entry.Blob.Size + FsshttpbSize.BinaryItem(entry.ClockData.Length));
            writer.WriteExtendedGuid(entry.Blob);
            writer.WriteBinaryItem(entry.ClockData.Span);
        }

        writer.WriteEnd(StreamObjectType.ContentTagKnowledge);
    }
}

/// <summary>
/// Version Token Knowledge ([MS-FSSHTTPB] 2.2.1.13.6): a token of the version of the file, opaque to the protocol.
/// </summary>
record VersionTokenKnowledge(ReadOnlyMemory<byte> Token) :
    SpecializedKnowledge
{
    public override Guid Kind => VersionTokenKnowledgeKind;

    public static VersionTokenKnowledge ReadData(ref FsshttpbReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.VersionTokenKnowledge);
        return new(fields.ReadRest());
    }

    protected override void WriteData(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.VersionTokenKnowledge, Token.Length);
        writer.WriteBytes(Token.Span);
    }
}

/// <summary>
/// A Specialized Knowledge of a kind this codec does not know, kept as its stream objects.
/// </summary>
record UnknownKnowledge(Guid UnknownKind, ReadOnlyMemory<byte> Data) :
    SpecializedKnowledge
{
    public override Guid Kind => UnknownKind;

    protected override void WriteData(ref FsshttpbWriter writer) =>
        writer.WriteBytes(Data.Span);
}
