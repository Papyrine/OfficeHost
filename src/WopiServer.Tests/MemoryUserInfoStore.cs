/// <summary>
/// The user info in memory.
/// </summary>
sealed class MemoryUserInfoStore : IWopiUserInfoStore
{
    ConcurrentDictionary<string, string> values = new();

    public Task<string?> Get(string userId, Cancel cancel = default) =>
        Task.FromResult(values.GetValueOrDefault(userId));

    public Task Set(string userId, string userInfo, Cancel cancel = default)
    {
        values[userId] = userInfo;
        return Task.CompletedTask;
    }
}