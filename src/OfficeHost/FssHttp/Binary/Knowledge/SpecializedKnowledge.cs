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

    public static SpecializedKnowledge Read(ref FssHttpBReader reader)
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

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.SpecializedKnowledge, 16);
        writer.WriteGuid(Kind);
        WriteData(ref writer);
        writer.WriteEnd(StreamObjectType.SpecializedKnowledge);
    }

    protected abstract void WriteData(ref FssHttpBWriter writer);
}