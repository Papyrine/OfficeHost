namespace WopiServer;

/// <summary>
/// Where the files are. File ids are chosen by the storage, and must not change when a file is saved or renamed, as
/// WOPI clients keep using them. Operations on a missing file throw a <see cref="FileNotFoundException"/>; other
/// failures are reported to the WOPI client as a server error.
/// </summary>
public interface IWopiStorage
{
    /// <summary>
    /// The file, or null when there is none with this id.
    /// </summary>
    Task<WopiFile?> GetFile(string fileId, Cancel cancel = default);

    /// <summary>
    /// The file of the given name in the folder of the given file, or null when there is none.
    /// </summary>
    Task<WopiFile?> GetSibling(string fileId, string name, Cancel cancel = default);

    /// <summary>
    /// Opens the content of the file. The caller disposes the stream.
    /// </summary>
    Task<Stream> OpenRead(string fileId, Cancel cancel = default);

    /// <summary>
    /// Replaces the content of the file, returning its new version.
    /// </summary>
    Task<WopiFile> Write(string fileId, Stream content, Cancel cancel = default);

    /// <summary>
    /// Creates a file of the given name in the folder of the given file. Returns null when a file of that name exists
    /// and <paramref name="overwrite"/> is false; replaces its content otherwise. The name is a legal file name.
    /// </summary>
    Task<WopiFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default);

    /// <summary>
    /// Renames the file, keeping its id. Returns null when a file of that name exists. The name is a legal file name.
    /// </summary>
    Task<WopiFile?> Rename(string fileId, string name, Cancel cancel = default);

    Task Delete(string fileId, Cancel cancel = default);
}

/// <summary>
/// A file of the <see cref="IWopiStorage"/>.
/// </summary>
/// <param name="Name">The name of the file, with its extension.</param>
/// <param name="OwnerId">A unique id of the owner of the file.</param>
/// <param name="Version">An id of the content of the file, which changes whenever the content does.</param>
public record WopiFile(string Id, string Name, long Size, string OwnerId, string Version, DateTimeOffset LastModified);

/// <summary>
/// Keeps the user info that WOPI clients store with PutUserInfo, such as which tips a user has seen. Without one,
/// PutUserInfo is not supported.
/// </summary>
public interface IWopiUserInfoStore
{
    /// <summary>
    /// The user info, or null when there is none.
    /// </summary>
    Task<string?> Get(string userId, Cancel cancel = default);

    Task Set(string userId, string userInfo, Cancel cancel = default);
}
