namespace OfficeHost;

/// <summary>
/// The files in an Azure Blob Storage container, shared by all the instances of the application. A file is the blob
/// <c>files/{id}</c>, with its folder, name and owner in its metadata, so a rename keeps the id and copies nothing. A
/// file is in a folder, a path the application chooses, and file names are unique in a folder, ignoring case: the
/// blob <c>names/{hash of the folder and name}</c> holds the id of the file of that name, created only when missing,
/// so two files never take the same name. A name left by a crash, naming a file that is gone or renamed, is taken
/// over. The version of a file changes only with its content.
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
        var file = await Upload(NewId(), folder, name, ownerId, content, cancel);
        if (await Claim(folder, name, file.Id, cancel) == null)
        {
            return file;
        }

        await FileBlob(file.Id).DeleteIfExistsAsync(cancellationToken: cancel);
        return null;
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
            file.Folder == folder &&
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

    public async Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        var metadata = new Dictionary<string, string>(source.Metadata)
        {
            [versionKey] = NewId()
        };
        try
        {
            // the metadata read, unless a rename changed it meanwhile
            await FileBlob(fileId).UploadAsync(
                content,
                new BlobUploadOptions
                {
                    Metadata = metadata,
                    Conditions = new()
                    {
                        IfMatch = source.ETag
                    }
                },
                cancel);
        }
        catch (RequestFailedException exception) when (exception.Status == 412 && content.CanSeek)
        {
            content.Position = 0;
            return await Write(fileId, content, cancel);
        }
        catch (RequestFailedException exception) when (exception.Status == 404)
        {
            throw new FileNotFoundException($"No file {fileId}", exception);
        }

        return (await ExistingProperties(fileId, cancel)).File;
    }

    public async Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        var (holder, _) = await ReadName(NameBlob(source.Folder, name), source.Folder, name, cancel);
        if (holder == null)
        {
            var file = await Upload(NewId(), source.Folder, name, source.File.OwnerId, content, cancel);
            holder = await Claim(source.Folder, name, file.Id, cancel);
            if (holder == null)
            {
                return file;
            }

            // created meanwhile
            await FileBlob(file.Id).DeleteIfExistsAsync(cancellationToken: cancel);
            if (content.CanSeek)
            {
                content.Position = 0;
            }
        }

        if (!overwrite)
        {
            return null;
        }

        return await Write(holder.File.Id, content, cancel);
    }

    public async Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default)
    {
        var source = await ExistingProperties(fileId, cancel);
        var oldName = source.File.Name;
        await SetName(fileId, name, cancel);
        // a change of case only is the same name
        if (SameName(oldName, name))
        {
            return (await ExistingProperties(fileId, cancel)).File;
        }

        // the file has the name before the name is claimed, so that a claim always names a file that has it
        if (await Claim(source.Folder, name, fileId, cancel) != null)
        {
            await SetName(fileId, oldName, cancel);
            return null;
        }

        await Release(source.Folder, oldName, fileId, cancel);
        return (await ExistingProperties(fileId, cancel)).File;
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
