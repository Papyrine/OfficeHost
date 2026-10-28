/// <summary>
/// Where the module sends its batches: the shadow (HttpSink), or a folder (<see cref="FolderSink"/>).
/// </summary>
interface ITapSink
{
    /// <param name="watermark">All exchanges started before it are in this batch or an earlier one, or tombstoned.</param>
    /// <param name="dropped">The exchanges dropped without a tombstone since the last batch.</param>
    /// <param name="final">The last batch of the app domain, which stops.</param>
    TapSent Send(IReadOnlyList<TapItem> items, DateTime watermark, long dropped, bool final);
}
