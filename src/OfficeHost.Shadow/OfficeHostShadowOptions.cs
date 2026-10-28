namespace OfficeHost;

/// <summary>
/// The options of the shadow (docs/shadow.md), bound from the <c>OfficeHostShadow</c> section by
/// <see cref="OfficeHostShadowExtensions.AddOfficeHostShadow"/>.
/// </summary>
public class OfficeHostShadowOptions
{
    /// <summary>
    /// The key the module in SharePoint sends (<c>X-OfficeHost-Shadow-Key</c>); exchanges without it are refused.
    /// Required.
    /// </summary>
    public string ApiKey { get; set; } = "";

    /// <summary>
    /// The user of OfficeHost of a user of SharePoint, by its login (<c>i:0#.w|contoso\alice</c>). By default its id is
    /// the whole login, so that users of two domains stay apart, and its name the account without its domain.
    /// </summary>
    public Func<string, ShadowUser> Users { get; set; } = ShadowUser.FromLogin;

    /// <summary>
    /// The level each outcome is logged at. "OfficeHost failed" and "the responses are overly different" are
    /// <see cref="ShadowOutcome.OfficeHostFailed"/> and <see cref="ShadowOutcome.Different"/>.
    /// </summary>
    public Dictionary<ShadowOutcome, LogLevel> Levels { get; set; } = new()
    {
        [ShadowOutcome.Same] = LogLevel.Trace,
        [ShadowOutcome.KnownDifference] = LogLevel.Debug,
        [ShadowOutcome.NotCompared] = LogLevel.Debug,
        [ShadowOutcome.Different] = LogLevel.Warning,
        [ShadowOutcome.OfficeHostFailed] = LogLevel.Error,
        [ShadowOutcome.ShadowFailed] = LogLevel.Warning,
        [ShadowOutcome.TapDropped] = LogLevel.Information
    };

    /// <summary>
    /// The differences that are expected, which make a <see cref="ShadowOutcome.KnownDifference"/>: by default those the
    /// lab found (<see cref="ShadowRule.Defaults"/>).
    /// </summary>
    public List<ShadowRule> KnownDifferences { get; set; } = [.. ShadowRule.Defaults];

    /// <summary>
    /// How long an exchange waits for those of the other SharePoint servers that started before it, when a server's
    /// watermark lags (a server whose sender backs off). Later ones make their document untrusted.
    /// </summary>
    public TimeSpan MaximumDelay { get; set; } = TimeSpan.FromSeconds(15);

    /// <summary>
    /// How long a SharePoint server (an instance of the module) counts as active without a batch or a heartbeat.
    /// </summary>
    public TimeSpan InstanceTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// How long the shadow keeps what it knows of a document without exchanges: longer than Office keeps a lock without
    /// refreshing it (43 minutes), so that a document is forgotten, and its next session starts afresh, only when no
    /// session of it can be open.
    /// </summary>
    public TimeSpan IdleTimeout { get; set; } = TimeSpan.FromHours(1);

    /// <summary>
    /// The most documents the shadow keeps what it knows of, two copies of each file among it: beyond it, those
    /// without exchanges for longest are forgotten, and start afresh at their next session.
    /// </summary>
    public int MaximumDocuments { get; set; } = 1000;

    /// <summary>
    /// The largest batch the ingest reads.
    /// </summary>
    public int MaxBatchMegabytes { get; set; } = 256;

    /// <summary>
    /// How often the counts of the outcomes are logged, at Information.
    /// </summary>
    public TimeSpan SummaryInterval { get; set; } = TimeSpan.FromMinutes(15);

    /// <summary>
    /// The most exchanges waiting to be replayed; beyond it, the ingest answers 503 and the module in SharePoint backs off
    /// and drops.
    /// </summary>
    public int QueueCapacity { get; set; } = 10_000;

    /// <summary>
    /// The percentage of documents the module in SharePoint sends, given to it with each answer; null to leave its own.
    /// </summary>
    public int? SamplePercent { get; set; }

    /// <summary>
    /// A folder to write the exchanges received to, in the format of the replay (docs/replay.md), a folder a day, to
    /// replay them again with <c>ReplayTests.Recording</c>. The exchanges hold the documents: keep it private.
    /// </summary>
    public string? RecordingFolder { get; set; }
}
