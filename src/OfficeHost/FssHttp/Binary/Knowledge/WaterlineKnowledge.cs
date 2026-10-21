/// <summary>
/// Waterline Knowledge ([MS-FSSHTTPB] 2.2.1.13.4).
/// </summary>
record WaterlineKnowledge(IReadOnlyList<WaterlineKnowledgeEntry> Entries) :
    SpecializedKnowledge
{
    public override Guid Kind => WaterlineKnowledgeKind;

    public static WaterlineKnowledge ReadData(ref FssHttpBReader reader)
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

    protected override void WriteData(ref FssHttpBWriter writer)
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