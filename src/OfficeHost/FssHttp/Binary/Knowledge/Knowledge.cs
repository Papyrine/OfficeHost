/// <summary>
/// Knowledge ([MS-FSSHTTPB] 2.2.1.13): what a client knows of the state of a file, as Specialized Knowledge of several
/// kinds.
/// </summary>
record Knowledge(IReadOnlyList<SpecializedKnowledge> Items)
{
    public static Knowledge Empty { get; } = new([]);

    public static Knowledge Read(ref FssHttpBReader reader)
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

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.Knowledge, 0);
        foreach (var item in Items)
        {
            item.Write(ref writer);
        }

        writer.WriteEnd(StreamObjectType.Knowledge);
    }
}