/// <summary>
/// Content Tag Knowledge ([MS-FSSHTTPB] 2.2.1.13.5).
/// </summary>
record ContentTagKnowledge(IReadOnlyList<ContentTagKnowledgeEntry> Entries) :
    SpecializedKnowledge
{
    public override Guid Kind => ContentTagKnowledgeKind;

    public static ContentTagKnowledge ReadData(ref FssHttpBReader reader)
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

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.ContentTagKnowledge, 0);
        foreach (var entry in Entries)
        {
            writer.WriteStart(StreamObjectType.ContentTagKnowledgeEntry, entry.Blob.Size + FssHttpBSize.BinaryItem(entry.ClockData.Length));
            writer.WriteExtendedGuid(entry.Blob);
            writer.WriteBinaryItem(entry.ClockData.Span);
        }

        writer.WriteEnd(StreamObjectType.ContentTagKnowledge);
    }
}