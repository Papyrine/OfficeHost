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
