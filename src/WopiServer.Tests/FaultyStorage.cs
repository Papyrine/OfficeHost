/// <summary>
/// A storage failing the operations a test picks, to reach the error handling of the WOPI endpoints.
/// </summary>
sealed class FaultyStorage(IWopiStorage inner) :
    ForwardingStorage(inner)
{
    /// <summary>
    /// Given an operation name (the <see cref="IWopiStorage"/> method), the exception to throw, or null to go on.
    /// </summary>
    public Func<string, Exception?> Fault { get; set; } = _ => null;

    void Check(string operation)
    {
        var exception = Fault(operation);
        if (exception != null)
        {
            throw exception;
        }
    }

    public override Task<WopiFile?> GetFile(string fileId, Cancel cancel = default)
    {
        Check(nameof(GetFile));
        return base.GetFile(fileId, cancel);
    }

    public override Task<WopiFile?> GetSibling(string fileId, string name, Cancel cancel = default)
    {
        Check(nameof(GetSibling));
        return base.GetSibling(fileId, name, cancel);
    }

    public override Task<Stream> OpenRead(string fileId, Cancel cancel = default)
    {
        Check(nameof(OpenRead));
        return base.OpenRead(fileId, cancel);
    }

    public override Task<WopiFile> Write(string fileId, Stream content, Cancel cancel = default)
    {
        Check(nameof(Write));
        return base.Write(fileId, content, cancel);
    }

    public override Task<WopiFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default)
    {
        Check(nameof(CreateSibling));
        return base.CreateSibling(fileId, name, content, overwrite, cancel);
    }

    public override Task<WopiFile?> Rename(string fileId, string name, Cancel cancel = default)
    {
        Check(nameof(Rename));
        return base.Rename(fileId, name, cancel);
    }

    public override Task Delete(string fileId, Cancel cancel = default)
    {
        Check(nameof(Delete));
        return base.Delete(fileId, cancel);
    }
}
