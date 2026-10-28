/// <summary>
/// The clock of OfficeHost in a shadow deployment: while an exchange is replayed, the time SharePoint received it, else
/// the system's. So coauthoring sessions and locks expire as they did on SharePoint, however late the shadow replays.
/// </summary>
sealed class ShadowClock : TimeProvider
{
    static AsyncLocal<DateTimeOffset?> current = new();

    public override DateTimeOffset GetUtcNow() => current.Value ?? System.GetUtcNow();

    /// <summary>
    /// Sets the time of the replay of an exchange, until the result is disposed.
    /// </summary>
    public static IDisposable At(DateTimeOffset? time)
    {
        var previous = current.Value;
        current.Value = time;
        return new Restore(previous);
    }

    sealed class Restore(DateTimeOffset? previous) : IDisposable
    {
        public void Dispose() => current.Value = previous;
    }
}
