namespace OfficeHost;

/// <summary>
/// What the shadow found of a sub-response of SharePoint replayed to OfficeHost (docs/shadow.md). Each is logged with
/// an event id of its value and the level <see cref="OfficeHostShadowOptions.Levels"/> gives it.
/// </summary>
public enum ShadowOutcome
{
    /// <summary>
    /// OfficeHost answered as SharePoint did.
    /// </summary>
    Same = 1,

    /// <summary>
    /// The answers differ only in ways a <see cref="ShadowRule"/> of <see cref="OfficeHostShadowOptions.KnownDifferences"/>
    /// expects.
    /// </summary>
    KnownDifference = 2,

    /// <summary>
    /// Not compared: what the answers depend on is not known to the shadow, as a document Office opened from its cache,
    /// or a document whose exchanges the shadow did not all receive.
    /// </summary>
    NotCompared = 3,

    /// <summary>
    /// The answers differ: OfficeHost answered otherwise than SharePoint.
    /// </summary>
    Different = 4,

    /// <summary>
    /// OfficeHost failed: it answered a server error, threw, or failed a sub-request that SharePoint did.
    /// </summary>
    OfficeHostFailed = 5,

    /// <summary>
    /// The shadow could not replay or compare an exchange, as a document whose content it could not get.
    /// </summary>
    ShadowFailed = 6,

    /// <summary>
    /// The module in SharePoint did not send exchanges: it dropped them (its queue was full) or sent a tombstone (too
    /// large, or the client went away).
    /// </summary>
    TapDropped = 7
}
