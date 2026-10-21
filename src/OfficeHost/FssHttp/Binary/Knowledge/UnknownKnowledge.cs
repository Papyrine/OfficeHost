/// <summary>
/// A Specialized Knowledge of a kind this codec does not know, kept as its stream objects.
/// </summary>
record UnknownKnowledge(Guid UnknownKind, ReadOnlyMemory<byte> Data) :
    SpecializedKnowledge
{
    public override Guid Kind => UnknownKind;

    protected override void WriteData(ref FssHttpBWriter writer) =>
        writer.WriteBytes(Data.Span);
}