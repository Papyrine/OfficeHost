/// <summary>
/// Forwards every call to another storage: tests override what they simulate.
/// </summary>
class ForwardingStorage(IWopiStorage inner) :
    IWopiStorage
{
    public virtual Task<WopiFile?> GetFile(string fileId, Cancel cancel = default) =>
        inner.GetFile(fileId, cancel);

    public virtual Task<WopiFile?> GetSibling(string fileId, string name, Cancel cancel = default) =>
        inner.GetSibling(fileId, name, cancel);

    public virtual Task<Stream> OpenRead(string fileId, Cancel cancel = default) =>
        inner.OpenRead(fileId, cancel);

    public virtual Task<WopiFile> Write(string fileId, Stream content, Cancel cancel = default) =>
        inner.Write(fileId, content, cancel);

    public virtual Task<WopiFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default) =>
        inner.CreateSibling(fileId, name, content, overwrite, cancel);

    public virtual Task<WopiFile?> Rename(string fileId, string name, Cancel cancel = default) =>
        inner.Rename(fileId, name, cancel);

    public virtual Task Delete(string fileId, Cancel cancel = default) =>
        inner.Delete(fileId, cancel);
}
