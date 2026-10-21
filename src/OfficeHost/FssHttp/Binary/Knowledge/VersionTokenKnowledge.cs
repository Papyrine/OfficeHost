/// <summary>
/// Version Token Knowledge ([MS-FSSHTTPB] 2.2.1.13.6): a token of the version of the file, opaque to the protocol.
/// </summary>
record VersionTokenKnowledge(ReadOnlyMemory<byte> Token) :
    SpecializedKnowledge
{
    public override Guid Kind => VersionTokenKnowledgeKind;

    public static VersionTokenKnowledge ReadData(ref FssHttpBReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.VersionTokenKnowledge);
        return new(fields.ReadRest());
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.VersionTokenKnowledge, Token.Length);
        writer.WriteBytes(Token.Span);
    }
}