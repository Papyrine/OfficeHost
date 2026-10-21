/// <summary>
/// A Query Access sub-request ([MS-FSSHTTPB] 2.2.2.1.2), which has no data.
/// </summary>
record QueryAccessRequest(ulong RequestId, ulong Priority = 0, Guid? TargetPartition = null) :
    SubRequest(RequestId, Priority, TargetPartition)
{
    public override RequestType Type => RequestType.QueryAccess;

    protected override void WriteData(ref FssHttpBWriter writer)
    {
    }
}