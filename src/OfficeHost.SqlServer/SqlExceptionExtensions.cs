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
    /// Runs an operation that sends the content again once, from where it started, when the content can be read again.
    /// </summary>
    public static async Task<T> Once<T>(Stream content, Func<Task<T>> operation)
    {
        long? start = null;
        if (content.CanSeek)
        {
            start = content.Position;
        }

        try
        {
            return await operation();
        }
        catch (SqlException exception) when (exception.IsDeadlock() && start != null)
        {
            content.Position = start.Value;
            return await operation();
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
