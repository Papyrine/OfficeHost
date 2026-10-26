static class SqlExceptionExtensions
{
    /// <summary>
    /// A violation of a primary key or of a unique index.
    /// </summary>
    public static bool IsDuplicateKey(this SqlException exception) =>
        exception.Number is 2627 or 2601;

    /// <summary>
    /// The statement was chosen as the victim of a deadlock, and rolled back.
    /// </summary>
    public static bool IsDeadlock(this SqlException exception) =>
        exception.Number == 1205;
}

/// <summary>
/// Runs an operation again once when it is the victim of a deadlock: SQL Server rolled it back, so running it again is
/// what SQL Server asks. Only for operations that can run twice, not those that send a content they read as they go.
/// </summary>
static class DeadlockRetry
{
    public static async Task<T> Once<T>(Func<Task<T>> operation)
    {
        try
        {
            return await operation();
        }
        catch (SqlException exception) when (exception.IsDeadlock())
        {
            return await operation();
        }
    }

    /// <summary>
    /// Runs an operation that sends the content again once, from where it started. A content that cannot be read again,
    /// as a request body, is copied to a temporary file first, for the run again to send it whole.
    /// </summary>
    public static async Task<T> Once<T>(Stream content, Func<Stream, Task<T>> operation, Cancel cancel)
    {
        if (content.CanSeek)
        {
            return await Seekable(content, operation);
        }

        await using var buffer = new FileStream(
            Path.GetTempFileName(),
            FileMode.Create,
            FileAccess.ReadWrite,
            FileShare.None,
            4096,
            FileOptions.Asynchronous | FileOptions.DeleteOnClose);
        await content.CopyToAsync(buffer, cancel);
        buffer.Position = 0;
        return await Seekable(buffer, operation);
    }

    static async Task<T> Seekable<T>(Stream content, Func<Stream, Task<T>> operation)
    {
        var start = content.Position;
        try
        {
            return await operation(content);
        }
        catch (SqlException exception) when (exception.IsDeadlock())
        {
            content.Position = start;
            return await operation(content);
        }
    }

    public static async Task Once(Func<Task> operation)
    {
        try
        {
            await operation();
        }
        catch (SqlException exception) when (exception.IsDeadlock())
        {
            await operation();
        }
    }
}
