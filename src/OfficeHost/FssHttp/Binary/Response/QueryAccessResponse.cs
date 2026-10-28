/// <summary>
/// A Query Access sub-response ([MS-FSSHTTPB] 2.2.3.1.1): whether reads and writes will succeed, as errors, the HRESULT
/// zero for yes.
/// </summary>
/// <param name="ExtraAccess">A third access response after the write access response, which the spec does not define:
/// SharePoint Server SE sends one (docs/fsshttp-notes.md). Kept as it was sent; OfficeHost sends none.</param>
record QueryAccessResponse(ulong RequestId, ResponseError ReadAccess, ResponseError WriteAccess, ResponseError? ExtraAccess = null) :
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
        ResponseError? extra = null;
        if (reader.PeekStart(StreamObjectType.ExtraAccessResponse))
        {
            reader.ReadStart(StreamObjectType.ExtraAccessResponse).EnsureEnd("an extra access response start");
            extra = ResponseError.Read(ref reader);
            reader.ReadEnd(StreamObjectType.ExtraAccessResponse);
        }

        return new(requestId, read, write, extra);
    }

    protected override void WriteData(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.ReadAccessResponse, 0);
        ReadAccess.Write(ref writer);
        writer.WriteEnd(StreamObjectType.ReadAccessResponse);
        writer.WriteStart(StreamObjectType.WriteAccessResponse, 0);
        WriteAccess.Write(ref writer);
        writer.WriteEnd(StreamObjectType.WriteAccessResponse);
        if (ExtraAccess != null)
        {
            writer.WriteStart(StreamObjectType.ExtraAccessResponse, 0);
            ExtraAccess.Write(ref writer);
            writer.WriteEnd(StreamObjectType.ExtraAccessResponse);
        }
    }
}
