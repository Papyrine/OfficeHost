using System;
using System.Collections.Generic;

/// <summary>
/// The answer of a sink to a batch: sent, or to retry after a while (the shadow is full or paused), or failed; and the
/// sampling the shadow asks for.
/// </summary>
sealed class TapSent
{
    public bool Sent { get; set; }

    public TimeSpan? RetryAfter { get; set; }

    public int? SamplePercent { get; set; }

    public string? Error { get; set; }
}

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
