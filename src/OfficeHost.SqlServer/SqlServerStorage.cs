namespace OfficeHost;

/// <summary>
/// The files in a SQL Server table, shared by all the instances of the application. A file is in a folder, a path the
/// application chooses, and file names are unique in a folder as the collation of the database compares them. The ids
/// of the files are chosen by <see cref="Add"/>, and the version of a file changes only with its content. Content is
/// read in chunks, each by a statement of its own, so that a slow download holds no connection nor lock; and sent
/// before a statement runs, so that a slow upload holds no lock either.
/// </summary>
public sealed class SqlServerStorage(string connectionString, string schema = "dbo") :
    IOfficeStorage
{
    string table = SqlServerTables.Name(schema, SqlServerTables.Files);

    const string columns = "Id, Name, Size, OwnerId, Version, LastModified";

    const string inserted = "inserted.Id, inserted.Name, inserted.Size, inserted.OwnerId, inserted.Version, inserted.LastModified";

    /// <summary>
    /// The longest folder, name and owner id, as the table holds them: longer ones are refused rather than cut.
    /// </summary>
    const int maxFolderLength = 400;

    const int maxNameLength = 255;

    const int maxIdLength = 450;

    /// <summary>
    /// The error a statement throws when the file it names is missing.
    /// </summary>
    const int noFile = 50001;

    /// <summary>
    /// Adds a file, returning it with its id. Returns null when the folder has a file of that name.
    /// </summary>
    /// <param name="folder">The folder of the file, of which siblings are the files of the same folder.</param>
    /// <exception cref="ArgumentException">The folder, the name or the owner id is longer than the table holds.</exception>
    public Task<OfficeFile?> Add(string folder, string name, string ownerId, Stream content, Cancel cancel = default) =>
        DeadlockRetry.Once(content, _ => AddAttempt(folder, name, ownerId, _, cancel), cancel);

    async Task<OfficeFile?> AddAttempt(string folder, string name, string ownerId, Stream content, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        var id = Base64Url.EncodeToString(RandomNumberGenerator.GetBytes(12));
        await using var command = Command(
            $"""
             insert into {table} (Id, Folder, Name, OwnerId, Content, Version, LastModified)
             output {inserted}
             values (@id, @folder, @name, @ownerId, @content, newid(), sysdatetimeoffset());
             """,
            connection,
            id);
        AddText(command, "@folder", folder, maxFolderLength);
        AddText(command, "@name", name, maxNameLength);
        AddText(command, "@ownerId", ownerId, maxIdLength);
        AddContent(command, content);
        try
        {
            return await ReadFile(command, cancel);
        }
        catch (SqlException exception) when (exception.IsDuplicateKey())
        {
            // a file of that name in the folder
            return null;
        }
    }

    async Task<SqlConnection> Open(Cancel cancel)
    {
        var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        return connection;
    }

    static SqlCommand Command(string sql, SqlConnection connection, string? fileId = null)
    {
        var command = new SqlCommand(sql, connection);
        if (fileId != null)
        {
            AddText(command, "@id", fileId, maxIdLength);
        }

        return command;
    }

    static void AddText(SqlCommand command, string name, string value, int maxLength)
    {
        // a longer value would be cut silently by the size of the parameter
        if (value.Length > maxLength)
        {
            throw new ArgumentException($"{name[1..]} is longer than {maxLength} characters", name[1..]);
        }

        command.Parameters.Add(name, SqlDbType.NVarChar, maxLength).Value = value;
    }

    static void AddContent(SqlCommand command, Stream content)
    {
        // a stream value is sent as it is read, before the statement runs, which a slow client can make take long
        command.Parameters.Add("@content", SqlDbType.VarBinary, -1).Value = content;
        command.CommandTimeout = 0;
    }

    static async Task<OfficeFile?> ReadFile(SqlCommand command, Cancel cancel)
    {
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            return null;
        }

        return ReadFile(reader);
    }

    static OfficeFile ReadFile(SqlDataReader reader) =>
        new(
            reader.GetString(0),
            reader.GetString(1),
            reader.GetInt64(2),
            reader.GetString(3),
            reader.GetGuid(4).ToString("N"),
            reader.GetDateTimeOffset(5));

    public Task<OfficeFile?> GetFile(string fileId, Cancel cancel = default) =>
        DeadlockRetry.Once(() => GetFileAttempt(fileId, cancel));

    async Task<OfficeFile?> GetFileAttempt(string fileId, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        await using var command = Command($"select {columns} from {table} where Id = @id", connection, fileId);
        return await ReadFile(command, cancel);
    }

    public Task<OfficeFile?> GetSibling(string fileId, string name, Cancel cancel = default) =>
        DeadlockRetry.Once(() => GetSiblingAttempt(fileId, name, cancel));

    async Task<OfficeFile?> GetSiblingAttempt(string fileId, string name, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        // the source with its sibling, if any: no row is no source
        await using var command = Command(
            $"""
             select f.Id, f.Name, f.Size, f.OwnerId, f.Version, f.LastModified
             from {table} s
             left join {table} f on f.Folder = s.Folder and f.Name = @name
             where s.Id = @id
             """,
            connection,
            fileId);
        AddText(command, "@name", name, maxNameLength);
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            throw new FileNotFoundException($"No file {fileId}");
        }

        if (await reader.IsDBNullAsync(0, cancel))
        {
            return null;
        }

        return ReadFile(reader);
    }

    public Task<Stream> OpenRead(string fileId, Cancel cancel = default) =>
        DeadlockRetry.Once(() => OpenReadAttempt(fileId, cancel));

    async Task<Stream> OpenReadAttempt(string fileId, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        await using var command = Command($"select Version, datalength(Content) from {table} where Id = @id", connection, fileId);
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            throw new FileNotFoundException($"No file {fileId}");
        }

        return new SqlServerReadStream(connectionString, table, fileId, reader.GetGuid(0), reader.GetInt64(1));
    }

    public Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default) =>
        DeadlockRetry.Once(content, _ => WriteAttempt(fileId, _, cancel), cancel);

    async Task<OfficeFile> WriteAttempt(string fileId, Stream content, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        // the file as written, not as another write may have left it since
        await using var command = Command(
            $"update {table} set Content = @content, Version = newid(), LastModified = sysdatetimeoffset() output {inserted} where Id = @id",
            connection,
            fileId);
        AddContent(command, content);
        return await ReadFile(command, cancel) ??
               throw new FileNotFoundException($"No file {fileId}");
    }

    /// <summary>
    /// One batch, whose content is sent before it runs: the lock on the name is held while the content is copied into
    /// the table only.
    /// </summary>
    public Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default) =>
        DeadlockRetry.Once(content, _ => CreateSiblingAttempt(fileId, name, _, overwrite, cancel), cancel);

    async Task<OfficeFile?> CreateSiblingAttempt(string fileId, string name, Stream content, bool overwrite, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        await using var command = Command(
            $"""
             set xact_abort on;
             begin transaction;
             declare @folder nvarchar({maxFolderLength}), @ownerId nvarchar({maxIdLength}), @existing nvarchar({maxIdLength});
             select @folder = Folder, @ownerId = OwnerId from {table} where Id = @id;
             if @folder is null
                 throw {noFile}, N'No file', 1;
             select @existing = Id from {table} with (updlock, holdlock) where Folder = @folder and Name = @name;
             if @existing is null
                 insert into {table} (Id, Folder, Name, OwnerId, Content, Version, LastModified)
                 output {inserted}
                 values (@newId, @folder, @name, @ownerId, @content, newid(), sysdatetimeoffset());
             else if @overwrite = 1
                 update {table} set Content = @content, Version = newid(), LastModified = sysdatetimeoffset()
                 output {inserted}
                 where Id = @existing;
             commit;
             """,
            connection,
            fileId);
        AddText(command, "@name", name, maxNameLength);
        command.Parameters.Add("@newId", SqlDbType.NVarChar, maxIdLength).Value = Base64Url.EncodeToString(RandomNumberGenerator.GetBytes(12));
        command.Parameters.Add("@overwrite", SqlDbType.Bit).Value = overwrite;
        AddContent(command, content);
        try
        {
            return await ReadFile(command, cancel);
        }
        catch (SqlException exception) when (exception.Number == noFile)
        {
            throw new FileNotFoundException($"No file {fileId}", exception);
        }
        catch (SqlException exception) when (exception.IsDuplicateKey())
        {
            // a file of that name in the folder
            return null;
        }
    }

    public Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default) =>
        DeadlockRetry.Once(() => RenameAttempt(fileId, name, cancel));

    async Task<OfficeFile?> RenameAttempt(string fileId, string name, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        await using var command = Command($"update {table} set Name = @name output {inserted} where Id = @id", connection, fileId);
        AddText(command, "@name", name, maxNameLength);
        try
        {
            return await ReadFile(command, cancel) ??
                   throw new FileNotFoundException($"No file {fileId}");
        }
        catch (SqlException exception) when (exception.IsDuplicateKey())
        {
            // a file of that name in the folder
            return null;
        }
    }

    public Task Delete(string fileId, Cancel cancel = default) =>
        DeadlockRetry.Once(() => DeleteAttempt(fileId, cancel));

    async Task DeleteAttempt(string fileId, Cancel cancel)
    {
        await using var connection = await Open(cancel);
        await using var command = Command($"delete from {table} where Id = @id", connection, fileId);
        if (await command.ExecuteNonQueryAsync(cancel) == 0)
        {
            throw new FileNotFoundException($"No file {fileId}");
        }
    }
}
