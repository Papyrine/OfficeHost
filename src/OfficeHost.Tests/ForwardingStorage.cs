/// <summary>
/// Forwards every call to another storage: tests override what they simulate.
/// </summary>
class ForwardingStorage(IOfficeStorage inner) :
    IOfficeStorage
{
    public virtual Task<OfficeFile?> GetFile(string fileId, Cancel cancel = default) =>
        inner.GetFile(fileId, cancel);

    public virtual Task<OfficeFile?> GetSibling(string fileId, string name, Cancel cancel = default) =>
        inner.GetSibling(fileId, name, cancel);

    public virtual Task<Stream> OpenRead(string fileId, Cancel cancel = default) =>
        inner.OpenRead(fileId, cancel);

    public virtual Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default) =>
        inner.Write(fileId, content, cancel);

    public virtual Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default) =>
        inner.CreateSibling(fileId, name, content, overwrite, cancel);

    public virtual Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default) =>
        inner.Rename(fileId, name, cancel);

    public virtual Task Delete(string fileId, Cancel cancel = default) =>
        inner.Delete(fileId, cancel);
}
