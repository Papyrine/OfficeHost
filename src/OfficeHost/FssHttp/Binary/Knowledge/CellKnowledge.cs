/// <summary>
/// Cell Knowledge ([MS-FSSHTTPB] 2.2.1.13.2): the data elements a client has, as ranges and single Serial Numbers.
/// </summary>
record CellKnowledge(IReadOnlyList<CellKnowledgeItem> Items) :
    SpecializedKnowledge
{
    public override Guid Kind => CellKnowledgeKind;

    public static CellKnowledge ReadData(ref FssHttpBReader reader)
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

    protected override void WriteData(ref FssHttpBWriter writer)
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