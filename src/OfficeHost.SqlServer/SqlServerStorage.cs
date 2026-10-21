namespace OfficeHost;

/// <summary>
/// The files in a SQL Server table, shared by all the instances of the application. A file is in a folder, a path the
/// application chooses, and file names are unique in a folder as the collation of the database compares them. The ids
/// of the files are chosen by <see cref="Add"/>, and the version of a file is its rowversion.
/// </summary>
public sealed class SqlServerStorage(string connectionString, string schema = "dbo") :
    IOfficeStorage
{
    string table = SqlServerTables.Name(schema, SqlServerTables.Files);

    const string columns = "Id, Name, Size, OwnerId, Version, LastModified";

    /// <summary>
    /// Adds a file, returning it with its id. Returns null when the folder has a file of that name.
    /// </summary>
    /// <param name="folder">The folder of the file, of which siblings are the files of the same folder.</param>
    public async Task<OfficeFile?> Add(string folder, string name, string ownerId, Stream content, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        return await Insert(connection, null, folder, name, ownerId, content, cancel);
    }

    async Task<SqlConnection> Open(Cancel cancel)
    {
        var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        return connection;
    }

    static SqlCommand Command(string sql, SqlConnection connection, SqlTransaction? transaction, string? fileId = null)
    {
        var command = new SqlCommand(sql, connection, transaction);
        if (fileId != null)
        {
            command.Parameters.Add("@id", SqlDbType.NVarChar, 450).Value = fileId;
        }

        return command;
    }

    static void AddContent(SqlCommand command, Stream content) =>
        // a stream value is sent as it is read
        command.Parameters.Add("@content", SqlDbType.VarBinary, -1).Value = content;

    static async Task<OfficeFile?> ReadFile(SqlCommand command, Cancel cancel)
    {
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            return null;
        }

        return new(
            reader.GetString(0),
            reader.GetString(1),
            reader.GetInt64(2),
            reader.GetString(3),
            Convert.ToHexString((byte[]) reader[4]),
            reader.GetDateTimeOffset(5));
    }

    async Task<OfficeFile?> Select(SqlConnection connection, SqlTransaction? transaction, string fileId, Cancel cancel)
    {
        await using var command = Command($"select {columns} from {table} where Id = @id", connection, transaction, fileId);
        return await ReadFile(command, cancel);
    }

    async Task<string> FolderOf(SqlConnection connection, SqlTransaction? transaction, string fileId, Cancel cancel)
    {
        await using var command = Command($"select Folder from {table} where Id = @id", connection, transaction, fileId);
        return (string?) await command.ExecuteScalarAsync(cancel) ??
               throw new FileNotFoundException($"No file {fileId}");
    }

    async Task<OfficeFile?> Insert(SqlConnection connection, SqlTransaction? transaction, string folder, string name, string ownerId, Stream content, Cancel cancel)
    {
        var id = Base64Url.EncodeToString(RandomNumberGenerator.GetBytes(12));
        await using var command = Command(
            $"""
             insert into {table} (Id, Folder, Name, OwnerId, Content, LastModified)
             values (@id, @folder, @name, @ownerId, @content, sysdatetimeoffset());
             """,
            connection,
            transaction,
            id);
        command.Parameters.Add("@folder", SqlDbType.NVarChar, 400).Value = folder;
        command.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = name;
        command.Parameters.Add("@ownerId", SqlDbType.NVarChar, 450).Value = ownerId;
        AddContent(command, content);
        try
        {
            await command.ExecuteNonQueryAsync(cancel);
        }
        catch (SqlException exception) when (exception.IsDuplicateKey())
        {
            // a file of that name in the folder
            return null;
        }

        return await Select(connection, transaction, id, cancel);
    }

    public async Task<OfficeFile?> GetFile(string fileId, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        return await Select(connection, null, fileId, cancel);
    }

    public async Task<OfficeFile?> GetSibling(string fileId, string name, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        var folder = await FolderOf(connection, null, fileId, cancel);
        await using var command = Command($"select {columns} from {table} where Folder = @folder and Name = @name", connection, null);
        command.Parameters.Add("@folder", SqlDbType.NVarChar, 400).Value = folder;
        command.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = name;
        return await ReadFile(command, cancel);
    }

    public async Task<Stream> OpenRead(string fileId, Cancel cancel = default)
    {
        var connection = await Open(cancel);
        try
        {
            await using var command = Command($"select Content from {table} where Id = @id", connection, null, fileId);
            var reader = await command.ExecuteReaderAsync(CommandBehavior.SequentialAccess | CommandBehavior.SingleRow, cancel);
            if (!await reader.ReadAsync(cancel))
            {
                await reader.DisposeAsync();
                throw new FileNotFoundException($"No file {fileId}");
            }

            return new SqlServerReadStream(connection, reader, reader.GetStream(0));
        }
        catch
        {
            await connection.DisposeAsync();
            throw;
        }
    }

    public async Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        await using var command = Command(
            $"""
             update {table} set Content = @content, LastModified = sysdatetimeoffset() where Id = @id;
             select {columns} from {table} where Id = @id;
             """,
            connection,
            null,
            fileId);
        AddContent(command, content);
        return await ReadFile(command, cancel) ??
               throw new FileNotFoundException($"No file {fileId}");
    }

    public async Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        await using var transaction = (SqlTransaction) await connection.BeginTransactionAsync(IsolationLevel.Serializable, cancel);
        string folder;
        string ownerId;
        await using (var source = Command($"select Folder, OwnerId from {table} where Id = @id", connection, transaction, fileId))
        await using (var reader = await source.ExecuteReaderAsync(cancel))
        {
            if (!await reader.ReadAsync(cancel))
            {
                throw new FileNotFoundException($"No file {fileId}");
            }

            folder = reader.GetString(0);
            ownerId = reader.GetString(1);
        }

        string? existing;
        await using (var find = Command($"select Id from {table} with (updlock) where Folder = @folder and Name = @name", connection, transaction))
        {
            find.Parameters.Add("@folder", SqlDbType.NVarChar, 400).Value = folder;
            find.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = name;
            existing = (string?) await find.ExecuteScalarAsync(cancel);
        }

        OfficeFile? file;
        if (existing == null)
        {
            file = await Insert(connection, transaction, folder, name, ownerId, content, cancel);
        }
        else
        {
            if (!overwrite)
            {
                return null;
            }

            await using var update = Command($"update {table} set Content = @content, LastModified = sysdatetimeoffset() where Id = @id", connection, transaction, existing);
            AddContent(update, content);
            await update.ExecuteNonQueryAsync(cancel);
            file = await Select(connection, transaction, existing, cancel);
        }

        await transaction.CommitAsync(cancel);
        return file;
    }

    public async Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        await using var command = Command(
            $"""
             update {table} set Name = @name where Id = @id;
             select {columns} from {table} where Id = @id;
             """,
            connection,
            null,
            fileId);
        command.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = name;
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

    public async Task Delete(string fileId, Cancel cancel = default)
    {
        await using var connection = await Open(cancel);
        await using var command = Command($"delete from {table} where Id = @id", connection, null, fileId);
        if (await command.ExecuteNonQueryAsync(cancel) == 0)
        {
            throw new FileNotFoundException($"No file {fileId}");
        }
    }
}
