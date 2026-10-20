namespace OfficeHost;

/// <summary>
/// The files of a folder, for development, tests and single-machine deployments. The ids of the files are kept in a
/// hidden <c>.wopi</c> folder at the root, so that they survive renames; <see cref="GetFileId"/> gives the id of a file
/// to open. The version of a file is its modification time, in ticks.
/// </summary>
public sealed class FileSystemStorage :
    IOfficeStorage
{
    const string metaFolderName = ".wopi";

    static StringComparison pathComparison =
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

    static JsonSerializerOptions indexOptions = new()
    {
        WriteIndented = true
    };

    string ownerId;
    string metaPath;
    string indexPath;
    Lock sync = new();

    // file id -> path relative to the root, with `/` separators
    Dictionary<string, string> paths;

    // path key -> file id
    Dictionary<string, string> ids = [];

    /// <param name="root">The folder of the files, which must exist.</param>
    /// <param name="ownerId">The OwnerId of all the files.</param>
    public FileSystemStorage(string root, string ownerId = "owner")
    {
        Root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));
        if (!Directory.Exists(Root))
        {
            throw new DirectoryNotFoundException($"No folder {Root}");
        }

        this.ownerId = ownerId;
        metaPath = Path.Combine(Root, metaFolderName);
        indexPath = Path.Combine(metaPath, "ids.json");
        paths = LoadIndex(indexPath);
        foreach (var (id, path) in paths)
        {
            ids[Key(path)] = id;
        }
    }

    /// <summary>
    /// The full path of the folder of the files.
    /// </summary>
    public string Root { get; }

    static Dictionary<string, string> LoadIndex(string path)
    {
        if (!File.Exists(path))
        {
            return [];
        }

        using var stream = File.OpenRead(path);
        return JsonSerializer.Deserialize<Dictionary<string, string>>(stream) ?? [];
    }

    static string Key(string relativePath)
    {
        if (pathComparison == StringComparison.OrdinalIgnoreCase)
        {
            return relativePath.ToUpperInvariant();
        }

        return relativePath;
    }

    /// <summary>
    /// The id of the file at the given path, relative to the root, assigned on first use.
    /// </summary>
    public string GetFileId(string relativePath)
    {
        var full = FullPath(relativePath);
        if (!File.Exists(full))
        {
            throw new FileNotFoundException($"No file {relativePath}", full);
        }

        return IdOf(full);
    }

    /// <summary>
    /// The full path of the file with the given id, or null when the id is unknown.
    /// </summary>
    public string? GetPath(string fileId)
    {
        lock (sync)
        {
            if (paths.TryGetValue(fileId, out var path))
            {
                return Path.GetFullPath(Path.Combine(Root, path));
            }

            return null;
        }
    }

    /// <summary>
    /// The full path of a path relative to the root, refusing those outside of it and the metadata folder.
    /// </summary>
    string FullPath(string relativePath)
    {
        var full = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.Join(Root, relativePath)));
        if (!IsBelow(full, Root) ||
            string.Equals(full, Root, pathComparison) ||
            IsBelow(full, metaPath))
        {
            throw new UnauthorizedAccessException($"{relativePath} is outside of the storage");
        }

        return full;
    }

    static bool IsBelow(string path, string folder) =>
        string.Equals(path, folder, pathComparison) ||
        path.StartsWith(folder + Path.DirectorySeparatorChar, pathComparison);

    string Relative(string full) =>
        Path.GetRelativePath(Root, full).Replace('\\', '/');

    string IdOf(string full)
    {
        var relative = Relative(full);
        lock (sync)
        {
            if (ids.TryGetValue(Key(relative), out var id))
            {
                return id;
            }

            id = Base64Url.EncodeToString(RandomNumberGenerator.GetBytes(12));
            paths[id] = relative;
            ids[Key(relative)] = id;
            SaveIndex();
            return id;
        }
    }

    void SaveIndex()
    {
        Directory.CreateDirectory(metaPath);
        var temp = $"{indexPath}.{Guid.NewGuid():N}.tmp";
        File.WriteAllBytes(temp, JsonSerializer.SerializeToUtf8Bytes(paths, indexOptions));
        File.Move(temp, indexPath, true);
    }

    string ExistingPath(string fileId) =>
        GetPath(fileId) ??
        throw new FileNotFoundException($"No file {fileId}");

    /// <summary>
    /// The path of a file named <paramref name="name"/> in the folder of the given file.
    /// </summary>
    string SiblingPath(string fileId, string name)
    {
        var folder = Path.GetDirectoryName(ExistingPath(fileId))!;
        var full = FullPath(Path.Join(Path.GetRelativePath(Root, folder), name));
        if (!string.Equals(Path.GetDirectoryName(full), folder, pathComparison))
        {
            throw new UnauthorizedAccessException($"{name} is not a file name");
        }

        return full;
    }

    /// <summary>
    /// The file, from a <see cref="FileInfo"/> that exists: its properties are read once, so that a file deleted
    /// meanwhile does not fail.
    /// </summary>
    OfficeFile Info(string fileId, FileInfo info)
    {
        var modified = info.LastWriteTimeUtc;
        return new(
            fileId,
            info.Name,
            info.Length,
            ownerId,
            modified.Ticks.ToString(CultureInfo.InvariantCulture),
            new(modified, TimeSpan.Zero));
    }

    public Task<OfficeFile?> GetFile(string fileId, Cancel cancel = default)
    {
        var full = GetPath(fileId);
        if (full == null)
        {
            return Task.FromResult<OfficeFile?>(null);
        }

        var info = new FileInfo(full);
        if (!info.Exists)
        {
            return Task.FromResult<OfficeFile?>(null);
        }

        return Task.FromResult<OfficeFile?>(Info(fileId, info));
    }

    public Task<OfficeFile?> GetSibling(string fileId, string name, Cancel cancel = default)
    {
        var info = new FileInfo(SiblingPath(fileId, name));
        if (!info.Exists)
        {
            return Task.FromResult<OfficeFile?>(null);
        }

        return Task.FromResult<OfficeFile?>(Info(IdOf(info.FullName), info));
    }

    public Task<Stream> OpenRead(string fileId, Cancel cancel = default)
    {
        var full = ExistingPath(fileId);
        try
        {
            return Task.FromResult<Stream>(new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 81920, FileOptions.Asynchronous));
        }
        catch (DirectoryNotFoundException exception)
        {
            throw new FileNotFoundException(exception.Message, full, exception);
        }
    }

    public async Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default)
    {
        var full = ExistingPath(fileId);
        if (!File.Exists(full))
        {
            throw new FileNotFoundException($"No file {fileId}", full);
        }

        var temp = await WriteTemp(content, cancel);
        Replace(temp, full);
        return Info(fileId, new(full));
    }

    public async Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default)
    {
        var full = SiblingPath(fileId, name);
        if (!overwrite &&
            File.Exists(full))
        {
            return null;
        }

        var temp = await WriteTemp(content, cancel);
        if (overwrite)
        {
            Replace(temp, full);
        }
        else
        {
            try
            {
                File.Move(temp, full, false);
            }
            catch (IOException) when (File.Exists(full))
            {
                // created meanwhile
                File.Delete(temp);
                return null;
            }
        }

        return Info(IdOf(full), new(full));
    }

    async Task<string> WriteTemp(Stream content, Cancel cancel)
    {
        var folder = Path.Combine(metaPath, "tmp");
        Directory.CreateDirectory(folder);
        var temp = Path.Combine(folder, Guid.NewGuid().ToString("N"));
        try
        {
            await using var target = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, FileOptions.Asynchronous);
            await content.CopyToAsync(target, cancel);
        }
        catch
        {
            File.Delete(temp);
            throw;
        }

        return temp;
    }

    /// <summary>
    /// Replaces the file with the temporary one, keeping the modification time increasing so that every write is a new
    /// version, even within the resolution of the file system's clock.
    /// </summary>
    void Replace(string temp, string full)
    {
        DateTime? previous = null;
        if (File.Exists(full))
        {
            previous = File.GetLastWriteTimeUtc(full);
        }

        try
        {
            Move(temp, full);
        }
        catch
        {
            File.Delete(temp);
            throw;
        }

        if (previous != null &&
            File.GetLastWriteTimeUtc(full) <= previous)
        {
            File.SetLastWriteTimeUtc(full, previous.Value.AddTicks(1));
        }
    }

    void Move(string temp, string full)
    {
        try
        {
            File.Move(temp, full, true);
        }
        catch (Exception exception) when (OperatingSystem.IsWindows() && exception is UnauthorizedAccessException or IOException)
        {
            // Windows does not replace a file that is open, as by a GetFile, but lets it be moved away and deleted
            var replaced = Path.Combine(metaPath, "tmp", Guid.NewGuid().ToString("N"));
            File.Move(full, replaced);
            try
            {
                File.Move(temp, full);
            }
            catch
            {
                // the file stays as it was, rather than left in the temporary folder
                File.Move(replaced, full);
                throw;
            }

            File.Delete(replaced);
        }
    }

    public Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default)
    {
        var full = ExistingPath(fileId);
        if (!File.Exists(full))
        {
            throw new FileNotFoundException($"No file {fileId}", full);
        }

        var target = SiblingPath(fileId, name);
        // a change of case only, on a case insensitive file system, is the same file
        var sameFile = string.Equals(full, target, pathComparison);
        if (!sameFile &&
            File.Exists(target))
        {
            return Task.FromResult<OfficeFile?>(null);
        }

        try
        {
            File.Move(full, target, false);
        }
        catch (IOException) when (!sameFile && File.Exists(target))
        {
            // created meanwhile
            return Task.FromResult<OfficeFile?>(null);
        }

        lock (sync)
        {
            ids.Remove(Key(Relative(full)));
            var relative = Relative(target);
            // a file deleted from the new path outside of the storage leaves its id, which would name this file too
            if (ids.TryGetValue(Key(relative), out var previous))
            {
                paths.Remove(previous);
            }

            paths[fileId] = relative;
            ids[Key(relative)] = fileId;
            SaveIndex();
        }

        return Task.FromResult<OfficeFile?>(Info(fileId, new(target)));
    }

    public Task Delete(string fileId, Cancel cancel = default)
    {
        var full = ExistingPath(fileId);
        if (!File.Exists(full))
        {
            throw new FileNotFoundException($"No file {fileId}", full);
        }

        File.Delete(full);
        lock (sync)
        {
            ids.Remove(Key(Relative(full)));
            paths.Remove(fileId);
            SaveIndex();
        }

        return Task.CompletedTask;
    }
}
