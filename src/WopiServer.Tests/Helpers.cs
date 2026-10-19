/// <summary>
/// A clock the tests move by hand, so lock expirations need no sleeping.
/// </summary>
class ManualTime : TimeProvider
{
    public DateTimeOffset Now { get; set; } = DateTimeOffset.UtcNow;

    public override DateTimeOffset GetUtcNow() =>
        Now;

    public void Advance(TimeSpan span) =>
        Now += span;
}

/// <summary>
/// A temporary folder, deleted on dispose.
/// </summary>
sealed class TempFolder : IDisposable
{
    public string Path { get; } = Directory.CreateTempSubdirectory("wopitest").FullName;

    public string Combine(params string[] parts) =>
        System.IO.Path.Combine([Path, .. parts]);

    public void Dispose()
    {
        try
        {
            Directory.Delete(Path, true);
        }
        catch (IOException)
        {
        }
    }
}

static class Helpers
{
    public static async Task<T> Throws<T>(Func<Task> action)
        where T : Exception
    {
        try
        {
            await action();
        }
        catch (T exception)
        {
            return exception;
        }

        throw new($"Expected {typeof(T).Name}");
    }

    public static Stream Stream(string content) =>
        new MemoryStream(Encoding.UTF8.GetBytes(content));
}
