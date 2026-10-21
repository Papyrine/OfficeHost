/// <summary>
/// The sub-response of a sub-request that failed (the status bit of 2.2.3.1), with its error.
/// </summary>
record FailedSubResponse(ulong RequestId, RequestType FailedType, ResponseError Error) :
    SubResponse(RequestId)
{
    public override RequestType Type => FailedType;

    protected override void WriteData(ref FssHttpBWriter writer) => Error.Write(ref writer);
}