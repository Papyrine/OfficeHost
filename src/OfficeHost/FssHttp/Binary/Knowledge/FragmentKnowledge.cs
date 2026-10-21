/// <summary>
/// Fragment Knowledge ([MS-FSSHTTPB] 2.2.1.13.3).
/// </summary>
record FragmentKnowledge(IReadOnlyList<FragmentKnowledgeEntry> Entries) :
    SpecializedKnowledge
{
    public override Guid Kind => FragmentKnowledgeKind;

    public static FragmentKnowledge ReadData(ref FssHttpBReader reader)
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

    protected override void WriteData(ref FssHttpBWriter writer)
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