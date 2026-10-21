/// <summary>
/// A Query Access sub-response ([MS-FSSHTTPB] 2.2.3.1.1): whether reads and writes will succeed, as errors, the HRESULT
/// zero for yes.
/// </summary>
record QueryAccessResponse(ulong RequestId, ResponseError ReadAccess, ResponseError WriteAccess) :
    SubResponse(RequestId)
{
    public override RequestType Type => RequestType.QueryAccess;

    public static QueryAccessResponse ReadData(ref FssHttpBReader reader, ulong requestId)
    {
        reader.ReadStart(StreamObjectType.ReadAccessResponse).EnsureEnd("a read access response start");
        var read = ResponseError.Read(ref reader);
        reader.ReadEnd(StreamObjectType.ReadAccessResponse);
        reader.ReadStart(StreamObjectType.WriteAccessResponse).EnsureEnd("a write access response start");
        var write = ResponseError.Read(ref reader);
        reader.ReadEnd(StreamObjectType.WriteAccessResponse);
        return new(requestId, read, write);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.ReadAccessResponse, 0);
        ReadAccess.Write(ref writer);
        writer.WriteEnd(StreamObjectType.ReadAccessResponse);
        writer.WriteStart(StreamObjectType.WriteAccessResponse, 0);
        WriteAccess.Write(ref writer);
        writer.WriteEnd(StreamObjectType.WriteAccessResponse);
    }
}