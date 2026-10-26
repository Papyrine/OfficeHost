namespace OfficeHost;

/// <summary>
/// The files in an Azure Blob Storage container, shared by all the instances of the application. A file is the blob
/// <c>files/{id}</c>, with its folder, name and owner in its metadata, so a rename keeps the id and copies nothing. A
/// file is in a folder, a path the application chooses, and file names are unique in a folder, ignoring case, as are
/// folders, as in a database of the default collation: the blob <c>names/{hash of the folder and name}</c> holds the id
/// of the file of that name, created only when missing, so two files never take the same name. A name left by a crash,
/// naming a file that is gone or renamed, is taken over. The version of a file changes only with its content.
/// </summary>
public sealed class BlobStorage(BlobContainerClient container) :
    IOfficeStorage
{
    const string folderKey = "folder";
    const string nameKey = "name";
    const string ownerKey = "owner";
    const string versionKey = "version";

    /// <summary>
    /// Adds a file, returning it with its id. Returns null when the folder has a file of that name.
    /// </summary>
    /// <param name="folder">The folder of the file, of which siblings are the files of the same folder.</param>
    public async Task<OfficeFile?> Add(string folder, string name, string ownerId, Stream content, Cancel cancel = default)
    {
        // a name taken is known before the content is uploaded
        var (holder, _) = await ReadName(NameBlob(folder, name), folder, name, cancel);
        if (holder != null)
        {
            return null;
        }

        var file = await Upload(NewId(), folder, name, ownerId, content, cancel);
        try
        {
            if (await Claim(folder, name, file.Id, cancel) == null)
            {
                return file;
            }
        }
        catch
        {
            await DeleteQuietly(file.Id);
            throw;
        }

        await DeleteQuietly(file.Id);
        return null;
    }

    /// <summary>
    /// Deletes the blob of a file that was uploaded but not named, whatever the request becomes.
    /// </summary>
    async Task DeleteQuietly(string fileId)
    {
        try
        {
            await FileBlob(fileId).DeleteIfExistsAsync(cancellationToken: Cancel.None);
        }
        catch (RequestFailedException)
        {
            // an unnamed blob is only unused
        }
    }

    static string NewId() =>
        Base64Url.EncodeToString(RandomNumberGenerator.GetBytes(12));

    BlobClient FileBlob(string fileId) =>
        container.GetBlobClient($"files/{fileId}");

    BlobClient NameBlob(string folder, string name)
    {
        var key = $"{folder}\n{name}".ToUpperInvariant();
        return container.GetBlobClient($"names/{Base64Url.EncodeToString(SHA256.HashData(Encoding.UTF8.GetBytes(key)))}");
    }

    // names and folders are compared as the name blob hashes them, upper-cased
    static bool SameName(string a, string b) =>
        string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    // metadata is ASCII only
    static string Encode(string value) =>
        Base64Url.EncodeToString(Encoding.UTF8.GetBytes(value));

    static string Decode(string value) =>
        Encoding.UTF8.GetString(Base64Url.DecodeFromChars(value));

    static Dictionary<string, string> Metadata(string folder, string name, string ownerId) =>
        new()
        {
            [folderKey] = Encode(folder),
            [nameKey] = Encode(name),
            [ownerKey] = Encode(ownerId),
            [versionKey] = NewId()
        };

    record Properties(OfficeFile File, string Folder, ETag ETag, IDictionary<string, string> Metadata);

    async Task<Properties?> GetProperties(string fileId, Cancel cancel)
    {
        try
        {
            BlobProperties properties = await FileBlob(fileId).GetPropertiesAsync(cancellationToken: cancel);
            var metadata = properties.Metadata;
            var file = new OfficeFile(
                fileId,
                Decode(metadata[nameKey]),
                properties.ContentLength,
                Decode(metadata[ownerKey]),
                metadata[versionKey],
                properties.LastModified);
            return new(file, Decode(metadata[folderKey]), properties.ETag, metadata);
        }
        catch (RequestFailedException exception) when (exception.Status == 404)
        {
            return null;
        }
    }

    async Task<Properties> ExistingProperties(string fileId, Cancel cancel) =>
        await GetProperties(fileId, cancel) ??
        throw new FileNotFoundException($"No file {fileId}");

    /// <summary>
    /// The file that has the name in the folder, or null when there is none, or when the name is left over.
    /// </summary>
    async Task<(Properties? File, ETag ETag)> ReadName(BlobClient nameBlob, string folder, string name, Cancel cancel)
    {
        BlobDownloadResult download;
        try
        {
            download = await nameBlob.DownloadContentAsync(cancel);
        }
        catch (RequestFailedException exception) when (exception.Status == 404)
        {
            return (null, default);
        }

        var file = await GetProperties(download.Content.ToString(), cancel);
        if (file != null &&
            SameName(file.Folder, folder) &&
            SameName(file.File.Name, name))
        {
            return (file, download.Details.ETag);
        }

        return (null, download.Details.ETag);
    }

    /// <summary>
    /// Gives the name in the folder to the file, which exists with that name. Returns the file that has the name, when
    /// another has it.
    /// </summary>
    async Task<Properties?> Claim(string folder, string name, string fileId, Cancel cancel)
    {
        var nameBlob = NameBlob(folder, name);
        while (true)
        {
            var (holder, etag) = await ReadName(nameBlob, folder, name, cancel);
            if (holder != null)
            {
                if (holder.File.Id == fileId)
                {
                    return null;
                }

                return holder;
            }

            var conditions = new BlobRequestConditions();
            if (etag == default)
            {
                conditions.IfNoneMatch = ETag.All;
            }
            else
            {
                conditions.IfMatch = etag;
            }

            try
            {
                await nameBlob.UploadAsync(
                    BinaryData.FromString(fileId),
                    new BlobUploadOptions
                    {
                        Conditions = conditions
                    },
                    cancel);
                return null;
            }
            catch (RequestFailedException exception) when (exception.Status is 409 or 412)
            {
                // claimed or changed meanwhile
            }
        }
    }

    /// <summary>
    /// Forgets the name in the folder, when it is still the file's.
    /// </summary>
    async Task Release(string folder, string name, string fileId, Cancel cancel)
    {
        var nameBlob = NameBlob(folder, name);
        try
        {
            var download = await nameBlob.DownloadContentAsync(cancel);
            if (download.Value.Content.ToString() != fileId)
            {
                return;
            }

            await nameBlob.DeleteAsync(
                conditions: new()
                {
                    IfMatch = download.Value.Details.ETag
                },
                cancellationToken: cancel);
        }
        catch (RequestFailedException exception) when (exception.Status is 404 or 412)
        {
            // gone, or taken meanwhile
        }
    }

    async Task<OfficeFile> Upload(string fileId, string folder, string name, string ownerId, Stream content, Cancel cancel)
    {
        await FileBlob(fileId).UploadAsync(
            content,
            new BlobUploadOptions
            {
                Metadata = Metadata(folder, name, ownerId),
                Conditions = new()
                {
                    IfNoneMatch = ETag.All
                }
            },
            cancel);
        return (await GetProperties(fileId, cancel))!.File;
    }

    public async Task<OfficeFile?> GetFile(string fileId, Cancel cancel = default) =>
        (await GetProperties(fileId, cancel))?.File;

    public async Task<OfficeFile?> GetSibling(string fileId, string name, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        var (sibling, _) = await ReadName(NameBlob(source.Folder, name), source.Folder, name, cancel);
        return sibling?.File;
    }

    public async Task<Stream> OpenRead(string fileId, Cancel cancel = default)
    {
        try
        {
            BlobDownloadStreamingResult download = await FileBlob(fileId).DownloadStreamingAsync(cancellationToken: cancel);
            return download.Content;
        }
        catch (RequestFailedException exception) when (exception.Status == 404)
        {
            throw new FileNotFoundException($"No file {fileId}", exception);
        }
    }

    /// <summary>
    /// Uploads the content with the metadata as last read, unless a rename changed it meanwhile: then it is uploaded
    /// again with the new one. A content that cannot be read twice, as a request body, is first kept in a temporary
    /// file.
    /// </summary>
    public async Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        await using var temporary = await Buffer(content, cancel);
        var buffered = temporary ?? content;
        var start = buffered.Position;
        var size = buffered.Length - start;
        var version = NewId();
        while (true)
        {
            var metadata = new Dictionary<string, string>(source.Metadata)
            {
                [versionKey] = version
            };
            try
            {
                BlobContentInfo info = await FileBlob(fileId).UploadAsync(
                    buffered,
                    new BlobUploadOptions
                    {
                        Metadata = metadata,
                        Conditions = new()
                        {
                            IfMatch = source.ETag
                        }
                    },
                    cancel);
                // the file as written, not as another write may have left it since
                return source.File with
                {
                    Size = size,
                    Version = version,
                    LastModified = info.LastModified
                };
            }
            catch (RequestFailedException exception) when (exception.Status == 412)
            {
                // renamed or written meanwhile
                buffered.Position = start;
                source = await ExistingProperties(fileId, cancel);
            }
            catch (RequestFailedException exception) when (exception.Status == 404)
            {
                throw new FileNotFoundException($"No file {fileId}", exception);
            }
        }
    }

    /// <summary>
    /// A copy of a content that cannot be read again, in a temporary file deleted when disposed; null for one that can.
    /// </summary>
    static async Task<FileStream?> Buffer(Stream content, Cancel cancel)
    {
        if (content.CanSeek)
        {
            return null;
        }

        var file = new FileStream(Path.GetTempFileName(), FileMode.Create, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.Asynchronous | FileOptions.DeleteOnClose);
        try
        {
            await content.CopyToAsync(file, cancel);
            file.Position = 0;
            return file;
        }
        catch
        {
            await file.DisposeAsync();
            throw;
        }
    }

    public async Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        var (holder, _) = await ReadName(NameBlob(source.Folder, name), source.Folder, name, cancel);
        if (holder != null)
        {
            if (!overwrite)
            {
                return null;
            }

            return await Write(holder.File.Id, content, cancel);
        }

        var file = await Upload(NewId(), source.Folder, name, source.File.OwnerId, content, cancel);
        var named = false;
        try
        {
            holder = await Claim(source.Folder, name, file.Id, cancel);
            if (holder == null)
            {
                named = true;
                return file;
            }

            // created meanwhile: the content was read, so the one uploaded is what replaces it
            if (overwrite)
            {
                await using var uploaded = await FileBlob(file.Id).OpenReadAsync(cancellationToken: cancel);
                return await Write(holder.File.Id, uploaded, cancel);
            }

            return null;
        }
        finally
        {
            if (!named)
            {
                await DeleteQuietly(file.Id);
            }
        }
    }

    public async Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        var oldName = source.File.Name;
        // a change of case only is the same name
        if (SameName(oldName, name))
        {
            await SetName(fileId, name, cancel);
            return (await ExistingProperties(fileId, cancel)).File;
        }

        // the file has the name before the name is claimed, so that a claim always names a file that has it
        await SetName(fileId, name, cancel);
        Properties? holder;
        try
        {
            holder = await Claim(source.Folder, name, fileId, cancel);
        }
        catch
        {
            await Restore(source.Folder, fileId, oldName);
            throw;
        }

        if (holder != null)
        {
            await Restore(source.Folder, fileId, oldName);
            return null;
        }

        await Release(source.Folder, oldName, fileId, cancel);
        return (await ExistingProperties(fileId, cancel)).File;
    }

    /// <summary>
    /// Gives a file its old name back after a rename failed, whatever the request becomes, and claims the name again, as
    /// while the file had the new one another file could take it.
    /// </summary>
    async Task Restore(string folder, string fileId, string oldName)
    {
        await SetName(fileId, oldName, Cancel.None);
        await Claim(folder, oldName, fileId, Cancel.None);
    }

    async Task SetName(string fileId, string name, Cancel cancel)
    {
        while (true)
        {
            var source = await ExistingProperties(fileId, cancel);
            var metadata = new Dictionary<string, string>(source.Metadata)
            {
                [nameKey] = Encode(name)
            };
            try
            {
                await FileBlob(fileId).SetMetadataAsync(
                    metadata,
                    new()
                    {
                        IfMatch = source.ETag
                    },
                    cancel);
                return;
            }
            catch (RequestFailedException exception) when (exception.Status == 412)
            {
                // written meanwhile
            }
            catch (RequestFailedException exception) when (exception.Status == 404)
            {
                throw new FileNotFoundException($"No file {fileId}", exception);
            }
        }
    }

    public async Task Delete(string fileId, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        try
        {
            await FileBlob(fileId).DeleteAsync(cancellationToken: cancel);
        }
        catch (RequestFailedException exception) when (exception.Status == 404)
        {
            throw new FileNotFoundException($"No file {fileId}", exception);
        }

        await Release(source.Folder, source.File.Name, fileId, cancel);
    }
}
